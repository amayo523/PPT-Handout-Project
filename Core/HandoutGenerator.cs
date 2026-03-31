using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using QuestPDF.Drawing;

namespace PptNotesHandoutMaker.Core
{
    public sealed class HandoutGenerator
    {
        private readonly HandoutOptions _opt;

        private readonly record struct NoteLine(int Level, string Text);

        public HandoutGenerator(HandoutOptions opt)
        {
            _opt = opt ?? throw new ArgumentNullException(nameof(opt));
        }

        public void Generate(string pptPath, string outputPdfPath, IProgress<string>? progress = null)
        {
            ValidateInputs(pptPath, outputPdfPath);

            Application? pptApp = null;
            Presentation? pres = null;

            bool powerPointWasAlreadyRunning = false;

            string tempImageDir = Path.Combine(
                Path.GetTempPath(),
                "ppt_handout_maker",
                Guid.NewGuid().ToString("N")
            );

            try
            {
                progress?.Report("Starting PowerPoint...");

                powerPointWasAlreadyRunning = PowerPointInterop.TryGetRunningPowerPoint(out pptApp);

                if (pptApp == null)
                {
                    powerPointWasAlreadyRunning = false;
                    pptApp = new Application();
                }

                pptApp.DisplayAlerts = PpAlertLevel.ppAlertsNone;

                progress?.Report("Opening presentation...");
                pres = pptApp.Presentations.Open(
                    FileName: pptPath,
                    ReadOnly: MsoTriState.msoTrue,
                    Untitled: MsoTriState.msoFalse,
                    WithWindow: MsoTriState.msoFalse
                );

                int slideCount = pres.Slides.Count;
                progress?.Report($"Slide count: {slideCount}");

                float slideRatioHW = pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth;
                progress?.Report($"Slide ratio (H/W): {slideRatioHW:0.0000}");

                var items = new List<(int SlideNumber, string SlideImagePath, List<NoteLine> NotesLines)>();

                for (int i = 1; i <= slideCount; i++)
                {
                    progress?.Report($"Exporting slide {i}/{slideCount}...");

                    Slide? slide = null;
                    try
                    {
                        slide = pres.Slides[i];

                        string imagePath = SlideExport.ExportSlidePng(
                            slide,
                            tempImageDir,
                            widthPx: _opt.SlideExportWidthPx);

                        var notesLines = NotesExtractor.GetSlideNotesLines(slide);

                        if (_opt.SkipSlidesWithNoNotes && notesLines.Count == 0)
                            continue;

                        items.Add((slide.SlideIndex, imagePath, notesLines));
                    }
                    finally
                    {
                        PowerPointInterop.FinalRelease(slide);
                    }
                }

                progress?.Report("Building PDF...");
                PdfBuilder.BuildHandoutPdf(_opt, items, outputPdfPath, slideRatioHW);
                progress?.Report("PDF built.");

                progress?.Report("Done.");
                progress?.Report($"PDF output: {outputPdfPath}");
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("PowerPoint COM error occurred.", ex);
            }
            finally
            {
                progress?.Report("Closing presentation...");
                if (pres != null)
                {
                    try { pres.Close(); } catch { }
                    PowerPointInterop.FinalRelease(pres);
                }

                progress?.Report("Releasing PowerPoint...");
                if (pptApp != null)
                {
                    try
                    {
                        // Only quit PowerPoint if WE started it and no other decks remain open
                        if (!powerPointWasAlreadyRunning)
                        {
                            int openCount = 0;
                            try { openCount = pptApp.Presentations.Count; } catch { }

                            if (openCount == 0)
                            {
                                try { pptApp.Quit(); } catch { }
                            }
                        }
                    }
                    finally
                    {
                        PowerPointInterop.FinalRelease(pptApp);
                    }
                }

                progress?.Report("Cleaning temp images...");
                try
                {
                    if (Directory.Exists(tempImageDir))
                        Directory.Delete(tempImageDir, recursive: true);
                }
                catch { }

                progress?.Report("Cleanup complete.");
            }
        }

        private static void ValidateInputs(string pptPath, string outputPdfPath)
        {
            if (string.IsNullOrWhiteSpace(pptPath))
                throw new ArgumentException("pptPath is required.", nameof(pptPath));

            if (!File.Exists(pptPath))
                throw new FileNotFoundException("PowerPoint file not found.", pptPath);

            if (string.IsNullOrWhiteSpace(outputPdfPath))
                throw new ArgumentException("outputPdfPath is required.", nameof(outputPdfPath));
        }

        // ============================================================
        // PowerPoint interop
        // ============================================================
        private static class PowerPointInterop
        {
            // GetActiveObject for COM (PowerPoint running instance)
            [DllImport("oleaut32.dll", PreserveSig = false)]
            private static extern void GetActiveObject(
                ref Guid rclsid,
                IntPtr reserved,
                [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

            public static bool TryGetRunningPowerPoint(out Application? app)
            {
                app = null;

                try
                {
                    // CLSID for PowerPoint.Application
                    Guid clsidPowerPoint = new("91493441-5A91-11CF-8700-00AA0060263B");

                    GetActiveObject(ref clsidPowerPoint, IntPtr.Zero, out object obj);
                    app = (Application)obj;
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            public static void FinalRelease(object? comObj)
            {
                if (comObj == null)
                    return;

                if (!OperatingSystem.IsWindows())
                    return;

                try { Marshal.FinalReleaseComObject(comObj); } catch { }
            }
        }

        // ============================================================
        // Slide export
        // ============================================================
        private static class SlideExport
        {
            public static string ExportSlidePng(Slide slide, string outDir, int widthPx = 1280)
            {
                Directory.CreateDirectory(outDir);

                string path = Path.Combine(outDir, $"slide_{slide.SlideIndex:D4}.png");

                var pres = slide.Parent as Presentation;
                float slideW = pres!.PageSetup.SlideWidth;
                float slideH = pres.PageSetup.SlideHeight;

                int heightPx = (int)Math.Round(widthPx * (slideH / slideW));

                slide.Export(path, "PNG", widthPx, heightPx);
                return path;
            }
        }

        // ============================================================
        // Notes extraction (nesting + numbering)
        // ============================================================
        private static class NotesExtractor
        {
            public static List<NoteLine> GetSlideNotesLines(Slide slide)
            {
                var raw = new List<(bool IsListItem, int IndentLevel, string Text)>();

                try
                {
                    var shapes = slide.NotesPage.Shapes;

                    // 1) Best path: Notes "body" placeholders first (most reliable)
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        var sh = shapes[i];
                        if (!IsNotesBodyPlaceholder(sh))
                            continue;

                        ExtractFromShapeTextFrame(sh, raw);
                    }

                    // 2) Fallback: scan all text shapes on NotesPage (excluding noise placeholders)
                    if (raw.Count == 0)
                    {
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            var sh = shapes[i];
                            if (IsNoiseNotesPlaceholder(sh))
                                continue;

                            ExtractFromShapeTextFrame(sh, raw);
                        }
                    }
                }
                catch
                {
                    // Slides without notes are normal — ignore
                }

                if (raw.Count == 0)
                    return new List<NoteLine>();

                // Level 0 = not bulleted, Level 1 = top-level bullet, Level 2+ nested
                var result = new List<NoteLine>(raw.Count);
                foreach (var p in raw)
                {
                    int level = p.IsListItem ? Math.Max(1, p.IndentLevel) : 0;
                    result.Add(new NoteLine(level, p.Text));
                }

                return result;
            }

            private static void ExtractFromShapeTextFrame(
                Microsoft.Office.Interop.PowerPoint.Shape sh,
                List<(bool IsListItem, int IndentLevel, string Text)> raw)
            {
                try
                {
                    if (sh.HasTextFrame != MsoTriState.msoTrue)
                        return;

                    var tf = sh.TextFrame;
                    if (tf == null || tf.HasText != MsoTriState.msoTrue)
                        return;

                    var tr = tf.TextRange;
                    if (tr == null)
                        return;

                    int paraCount;
                    try
                    {
                        paraCount = tr.Paragraphs().Count;
                    }
                    catch
                    {
                        return;
                    }

                    // Numbering counters per indent level for THIS shape
                    var numberCountersByIndent = new Dictionary<int, int>();

                    for (int p = 1; p <= paraCount; p++)
                    {
                        TextRange para;
                        try
                        {
                            para = tr.Paragraphs(p, 1);
                        }
                        catch
                        {
                            continue;
                        }

                        string text = (para.Text ?? "").Replace("\r", "").TrimEnd();
                        if (string.IsNullOrWhiteSpace(text))
                            continue;

                        int indentLevel = 1;
                        try { indentLevel = para.IndentLevel; } catch { indentLevel = 1; }

                        bool isNumbered = IsNumberedParagraph(para);
                        bool isBulleted = IsBulletedParagraph(para);

                        string prefix = "";

                        if (isNumbered)
                        {
                            // Reset deeper levels when numbering at a higher level
                            foreach (var k in numberCountersByIndent.Keys.Where(k => k > indentLevel).ToList())
                                numberCountersByIndent.Remove(k);

                            int current;
                            if (!numberCountersByIndent.ContainsKey(indentLevel))
                            {
                                current = GetBulletStartValueOrDefault(para, 1);
                                numberCountersByIndent[indentLevel] = current;
                            }
                            else
                            {
                                current = numberCountersByIndent[indentLevel] + 1;
                                numberCountersByIndent[indentLevel] = current;
                            }

                            object? styleObj = null;
                            try { styleObj = para.ParagraphFormat.Bullet.Style; } catch { }

                            prefix = FormatNumberLabel(current, styleObj) + " ";
                        }
                        else if (isBulleted)
                        {
                            // Bullets should not continue numbering runs at same indent
                            foreach (var k in numberCountersByIndent.Keys.Where(k => k >= indentLevel).ToList())
                                numberCountersByIndent.Remove(k);

                            if (!StartsWithBullet(text))
                                prefix = "• ";
                        }
                        else
                        {
                            // Plain paragraphs break numbering sequences at/under this indent
                            foreach (var k in numberCountersByIndent.Keys.Where(k => k >= indentLevel).ToList())
                                numberCountersByIndent.Remove(k);
                        }

                        bool isListItem = isNumbered || isBulleted;
                        raw.Add((isListItem, indentLevel, prefix + text));
                    }
                }
                catch
                {
                    // ignore individual shape issues
                }
            }

            private static bool IsNotesBodyPlaceholder(Microsoft.Office.Interop.PowerPoint.Shape sh)
            {
                try
                {
                    if (sh.Type != MsoShapeType.msoPlaceholder)
                        return false;

                    return sh.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody;
                }
                catch
                {
                    return false;
                }
            }

            private static bool IsNoiseNotesPlaceholder(Microsoft.Office.Interop.PowerPoint.Shape sh)
            {
                try
                {
                    if (sh.Type != MsoShapeType.msoPlaceholder)
                        return false;

                    var t = sh.PlaceholderFormat.Type;

                    return t == PpPlaceholderType.ppPlaceholderSlideNumber
                        || t == PpPlaceholderType.ppPlaceholderDate
                        || t == PpPlaceholderType.ppPlaceholderFooter
                        || t == PpPlaceholderType.ppPlaceholderHeader;
                }
                catch
                {
                    return false;
                }
            }

            private static bool StartsWithBullet(string s)
            {
                if (string.IsNullOrEmpty(s)) return false;

                char c = s[0];
                return c == '•' || c == '◦' || c == '▪' || c == '–' || c == '-' || c == '·';
            }

            private static bool IsNumberedParagraph(TextRange para)
            {
                try
                {
                    if (para.ParagraphFormat.Bullet.Visible != MsoTriState.msoTrue)
                        return false;

                    return para.ParagraphFormat.Bullet.Type == PpBulletType.ppBulletNumbered;
                }
                catch
                {
                    return false;
                }
            }

            private static bool IsBulletedParagraph(TextRange para)
            {
                try
                {
                    if (para.ParagraphFormat.Bullet.Visible != MsoTriState.msoTrue)
                        return false;

                    return para.ParagraphFormat.Bullet.Type == PpBulletType.ppBulletUnnumbered;
                }
                catch
                {
                    return false;
                }
            }

            private static int GetBulletStartValueOrDefault(TextRange para, int fallback = 1)
            {
                try
                {
                    return para.ParagraphFormat.Bullet.StartValue;
                }
                catch
                {
                    return fallback;
                }
            }

            private static string FormatNumberLabel(int value, object? bulletStyleObj)
            {
                string? styleName = null;

                try
                {
                    if (bulletStyleObj != null)
                    {
                        var t = bulletStyleObj.GetType();
                        if (t.IsEnum)
                            styleName = Enum.GetName(t, bulletStyleObj);
                    }
                }
                catch { }

                if (string.IsNullOrWhiteSpace(styleName))
                    return $"{value}.";

                return styleName switch
                {
                    "ppBulletArabicPeriod" => $"{value}.",
                    "ppBulletArabicParenRight" => $"{value})",
                    "ppBulletArabicParenBoth" => $"({value})",

                    "ppBulletAlphaUCPeriod" => $"{ToAlpha(value, upper: true)}.",
                    "ppBulletAlphaLCPeriod" => $"{ToAlpha(value, upper: false)}.",
                    "ppBulletAlphaUCParenRight" => $"{ToAlpha(value, upper: true)})",
                    "ppBulletAlphaLCParenRight" => $"{ToAlpha(value, upper: false)})",

                    "ppBulletRomanUCPeriod" => $"{ToRoman(value, upper: true)}.",
                    "ppBulletRomanLCPeriod" => $"{ToRoman(value, upper: false)}.",

                    _ => $"{value}."
                };
            }

            private static string ToAlpha(int value, bool upper)
            {
                if (value < 1) value = 1;

                string s = "";
                int n = value;

                while (n > 0)
                {
                    n--;
                    s = (char)('A' + (n % 26)) + s;
                    n /= 26;
                }

                return upper ? s : s.ToLowerInvariant();
            }

            private static string ToRoman(int value, bool upper)
            {
                if (value < 1) value = 1;

                var map = new (int v, string s)[]
                {
                    (1000,"M"),(900,"CM"),(500,"D"),(400,"CD"),
                    (100,"C"),(90,"XC"),(50,"L"),(40,"XL"),
                    (10,"X"),(9,"IX"),(5,"V"),(4,"IV"),(1,"I")
                };

                int n = value;
                string result = "";

                foreach (var (v, s) in map)
                {
                    while (n >= v)
                    {
                        result += s;
                        n -= v;
                    }
                }

                return upper ? result : result.ToLowerInvariant();
            }
        }

        // ============================================================
        // PDF generation (QuestPDF)
        // ============================================================
        private static class PdfBuilder
        {
            public static void BuildHandoutPdf(
                HandoutOptions opt,
                List<(int SlideNumber, string SlideImagePath, List<NoteLine> NotesLines)> items,
                string outputPdfPath,
                float slideRatioHW)
            {
                FontManager.RegisterFont(File.OpenRead("C:\\Windows\\Fonts\\arial.ttf"));
                FontManager.RegisterFont(File.OpenRead("C:\\Windows\\Fonts\\arialbd.ttf"));
                QuestPDF.Settings.License = LicenseType.Community;

                Document.Create(container =>
                {
                    container.Page(page =>
                    {
                        page.Size(PageSizes.Letter); // portrait
                        page.Margin(48);
                        page.DefaultTextStyle(x => x.FontFamily("Arial"));
                        page.DefaultTextStyle(x => x.FontSize(9));

                        // Thumbnail sizing constants
                        const float thumbBorder = 1f;
                        const float thumbPadding = 2f;
                        const float cellBorder = 1f;

                        // Header (pages 2+ only)
                        page.Header().ShowIf(ctx => ctx.PageNumber > 1).Element(h =>
                        {
                            string className = (opt.ClassName ?? "").Trim();

                            h.PaddingBottom(8)
                             .AlignRight()
                             .Text(className)
                             .FontSize(9)
                             .Bold();
                        });

                        // Footer (pages 2+ only)
                        page.Footer().ShowIf(ctx => ctx.PageNumber > 1).Element(f =>
                        {
                            string pdfTitle = (opt.PdfTitle ?? "").Trim();

                            f.PaddingTop(6)
                             .Row(row =>
                             {
                                 row.RelativeItem()
                                    .AlignLeft()
                                    .Text(pdfTitle)
                                    .FontSize(9)
                                    .Bold();

                                 row.RelativeItem()
                                    .AlignRight()
                                    .Text(t =>
                                    {
                                        t.DefaultTextStyle(x => x.FontSize(9).Bold());
                                        t.Span("Page ");
                                        t.CurrentPageNumber();
                                    });
                             });
                        });

                        page.Content()
                            .PaddingTop(12)
                            .Column(col =>
                            {
                                string className = (opt.ClassName ?? "").Trim();
                                string pdfTitle = (opt.PdfTitle ?? "").Trim();

                                col.Item().ShowIf(ctx => ctx.PageNumber == 1).PaddingBottom(18).Element(titleBlock =>
                                {
                                    titleBlock.AlignCenter().Column(tc =>
                                    {
                                        if (!string.IsNullOrWhiteSpace(className))
                                        {
                                            tc.Item()
                                              .Text(className)
                                              .FontSize(18)
                                              .Bold()
                                              .AlignCenter();
                                        }

                                        if (!string.IsNullOrWhiteSpace(pdfTitle))
                                        {
                                            tc.Item()
                                              .PaddingTop(4)
                                              .Text(pdfTitle)
                                              .FontSize(20)
                                              .Bold()
                                              .AlignCenter();
                                        }

                                        tc.Item()
                                          .PaddingTop(4)
                                          .Text("Instructor Guide")
                                          .FontSize(18)
                                          .Bold()
                                          .AlignCenter();
                                    });
                                });
                                for (int i = 0; i < items.Count; i++)
                                {
                                    var (slideNumber, imgPath, notesLines) = items[i];
                                    bool hasNotes = notesLines != null && notesLines.Count > 0;

                                    // Base thumbnail width
                                    const float baseThumbWidth = 240f;

                                    // Scale width if no notes
                                    float effectiveThumbWidth = hasNotes ? baseThumbWidth : baseThumbWidth * 0.7f;

                                    // Recalculate height from width
                                    float imageHeight = effectiveThumbWidth * slideRatioHW;

                                    // Recalculate box heights
                                    float thumbBoxHeight =
                                        imageHeight
                                        + (2 * thumbPadding)
                                        + (2 * thumbBorder)
                                        - cellBorder;

                                    float compactThumbBoxHeight = thumbBoxHeight * 0.45f;

                                    // Final height used everywhere
                                    float effectiveHeight = hasNotes ? thumbBoxHeight : compactThumbBoxHeight;

                                    col.Item()
                                       .ShowEntire()
                                       .PaddingVertical(hasNotes ? 20 : 6)
                                       .Element(outer =>
                                       {
                                           var cell = outer
                                               .Border(0)
                                               .BorderColor(Colors.Grey.Darken1)
                                               .Padding(0);

                                           float effectiveHeight = hasNotes ? thumbBoxHeight : compactThumbBoxHeight;

                                           cell = hasNotes
                                               ? cell.MinHeight(effectiveHeight)
                                               : cell.Height(effectiveHeight);

                                           cell.Element(block =>
                                           {
                                               block.Row(row =>
                                               {
                                                   row.RelativeItem(1.0f).Element(left =>
                                                   {
                                                       const float slideNumberRowHeight = 14f;
                                                       float thumbAreaHeight = Math.Max(1f, effectiveHeight - slideNumberRowHeight);

                                                       left.Height(effectiveHeight)
                                                           .AlignTop()
                                                           .Column(lc =>
                                                           {
                                                               lc.Spacing(0);

                                                               lc.Item()
                                                                 .Height(slideNumberRowHeight)
                                                                 .PaddingHorizontal(4)
                                                                 .AlignLeft()
                                                                 .AlignMiddle()
                                                                 .Text($"Slide {slideNumber}")
                                                                 .FontSize(8)
                                                                 .Bold();

                                                               lc.Item()
                                                                 .Height(thumbAreaHeight)
                                                                 .AlignCenter()
                                                                 .AlignMiddle()
                                                                 .Border(1)
                                                                 .BorderColor(Colors.Black)
                                                                 .Padding(2)
                                                                 .Image(imgPath)
                                                                 .FitArea();
                                                           });
                                                   });

                                                   row.ConstantItem(12);

                                                   row.RelativeItem(1.2f).Element(right =>
                                                   {
                                                       right.Column(ncol =>
                                                       {
                                                           ncol.Spacing(2);

                                                           var safeNotesLines = notesLines ?? new List<NoteLine>();

                                                           if (safeNotesLines.Count == 0)
                                                           {
                                                               ncol.Item().Text("(No notes)");
                                                               return;
                                                           }

                                                           const float indentStep = 14f;
                                                           float gutterWidth = ComputeGutterWidthForSlide(safeNotesLines);

                                                           foreach (var nl in safeNotesLines)
                                                           {
                                                               string line = string.IsNullOrWhiteSpace(nl.Text) ? " " : nl.Text;

                                                               int indentLevels = Math.Max(0, nl.Level - 1);
                                                               float leftPad = indentLevels * indentStep;

                                                               if (nl.Level <= 0)
                                                               {
                                                                   ncol.Item()
                                                                       .PaddingLeft(leftPad)
                                                                       .Text(line);
                                                                   continue;
                                                               }

                                                               var (prefix, body) = SplitListPrefix(line);

                                                               if (string.IsNullOrWhiteSpace(prefix))
                                                               {
                                                                   ncol.Item()
                                                                       .PaddingLeft(leftPad)
                                                                       .Text(body);
                                                                   continue;
                                                               }

                                                               ncol.Item()
                                                                   .PaddingLeft(leftPad)
                                                                   .Row(r =>
                                                                   {
                                                                       r.ConstantItem(gutterWidth)
                                                                        .AlignTop()
                                                                        .Text(prefix);

                                                                       r.RelativeItem()
                                                                        .AlignTop()
                                                                        .Text(body);
                                                                   });
                                                           }

                                                       });
                                                   });
                                               });
                                           });
                                       });

                                    if (i < items.Count - 1)
                                    {
                                        col.Item()
                                           .LineHorizontal(1)
                                           .LineColor(Colors.Grey.Lighten2);
                                    }
                                }
                            });
                    });
                })
                .GeneratePdf(outputPdfPath);
            }

            private static float ComputeGutterWidthForSlide(List<NoteLine>? notesLines)
            {
                const float baseWidth = 5f;
                const float perChar = 4.0f;
                const float minWidth = 12f;
                const float maxWidth = 32f;

                if (notesLines == null || notesLines.Count == 0)
                    return minWidth;

                int maxPrefixLen = 0;

                foreach (var nl in notesLines)
                {
                    if (nl.Level <= 0)
                        continue;

                    var (prefix, _) = SplitListPrefix(nl.Text ?? "");
                    if (!string.IsNullOrWhiteSpace(prefix))
                        maxPrefixLen = Math.Max(maxPrefixLen, prefix.Length);
                }

                float width = baseWidth + (maxPrefixLen * perChar);

                if (width < minWidth) width = minWidth;
                if (width > maxWidth) width = maxWidth;

                return width;
            }

            private static (string Prefix, string Body) SplitListPrefix(string line)
            {
                if (string.IsNullOrWhiteSpace(line))
                    return ("", " ");

                var s = line.TrimEnd();

                if (s.StartsWith("• "))
                    return ("•", s.Substring(2));

                int spaceIdx = s.IndexOf(' ');
                if (spaceIdx > 0 && spaceIdx <= 6)
                {
                    string firstToken = s.Substring(0, spaceIdx);
                    string rest = s.Substring(spaceIdx + 1);

                    bool looksNumbered =
                        firstToken.EndsWith(".") ||
                        firstToken.EndsWith(")") ||
                        (firstToken.StartsWith("(") && firstToken.EndsWith(")"));

                    if (looksNumbered)
                        return (firstToken, rest);
                }

                return ("", s);
            }
        }

        // ============================================================
        // Helpers currently NOT used by this class
        // Keeping them here so you don’t lose them, but clearly marked.
        // ============================================================

        // Output file naming (unused here)
        private static string GetNextAvailableFilePath(string desiredPath)
        {
            if (!File.Exists(desiredPath))
                return desiredPath;

            string directory = Path.GetDirectoryName(desiredPath)!;
            string baseName = Path.GetFileNameWithoutExtension(desiredPath);
            string extension = Path.GetExtension(desiredPath);

            int counter = 1;
            while (true)
            {
                string candidate = Path.Combine(directory, $"{baseName} ({counter}){extension}");
                if (!File.Exists(candidate))
                    return candidate;
                counter++;
            }
        }

        // Layout helper (unused here)
        private static float ComputeThumbBoxHeightFromLayout(
            float slideRatioHW,
            float pageWidthPts,
            float pageMarginPts,
            float gapPts,
            float leftRel,
            float rightRel,
            float thumbBorderPts,
            float thumbPaddingPts,
            float cellBorderPts)
        {
            float contentWidth = pageWidthPts - (2 * pageMarginPts);
            float columnAreaWidth = contentWidth - gapPts;

            float leftColWidth = columnAreaWidth * (leftRel / (leftRel + rightRel));

            float imageWidth = leftColWidth - (2 * thumbBorderPts) - (2 * thumbPaddingPts);
            if (imageWidth < 10f) imageWidth = 10f;

            float imageHeight = imageWidth * slideRatioHW;

            float boxHeight = imageHeight + (2 * thumbPaddingPts) + (2 * thumbBorderPts);

            // Adjust for the outer cell border so outlines end at the same visual height
            boxHeight -= cellBorderPts;

            return boxHeight;
        }

        // Reflection-based helpers (unused here)
        private static float InferIndentStep(List<float> indents)
        {
            if (indents == null || indents.Count < 2)
                return 0f;

            var distinct = indents.Distinct().OrderBy(x => x).ToList();
            if (distinct.Count < 2)
                return 0f;

            float step = float.MaxValue;
            for (int i = 1; i < distinct.Count; i++)
            {
                float diff = distinct[i] - distinct[i - 1];
                if (diff > 0.01f)
                    step = Math.Min(step, diff);
            }

            if (step == float.MaxValue || step < 0.5f)
                return 0f;

            return step;
        }

        private static float TryGetParagraphIndent(dynamic para2)
        {
            try
            {
                object? pf = GetProp(para2, "ParagraphFormat");
                if (pf == null) return 0f;

                object? left = GetProp(pf, "LeftIndent");
                if (left != null) return Convert.ToSingle(left);

                object? first = GetProp(pf, "FirstLineIndent");
                if (first != null) return Convert.ToSingle(first);
            }
            catch { }

            return 0f;
        }

        private static bool TryGetBulletVisible(dynamic para2)
        {
            try
            {
                object? pf = GetProp(para2, "ParagraphFormat");
                if (pf == null) return false;

                object? bullet = GetProp(pf, "Bullet");
                if (bullet == null) return false;

                object? visible = GetProp(bullet, "Visible");
                if (visible == null) return false;

                int v = Convert.ToInt32(visible); // MsoTriState: msoTrue = -1
                return v == -1;
            }
            catch
            {
                return false;
            }
        }

        private static object? GetProp(object obj, string name)
        {
            try
            {
                return obj.GetType().GetProperty(name)?.GetValue(obj, null);
            }
            catch
            {
                return null;
            }
        }
    }
}
