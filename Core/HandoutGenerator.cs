using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptNotesHandoutMaker.Core
{
    public sealed class HandoutGenerator
    {
        // Main flow:
        // 1. Validate inputs
        // 2. Open PowerPoint presentation
        // 3. Export slide images and extract notes
        // 4. Build the final PDF
        // 5. Clean up COM objects and temp files

        private readonly HandoutOptions _opt;

        private readonly record struct NoteLine(int Level, string Text);

        public HandoutGenerator(HandoutOptions opt)
        {
            _opt = opt ?? throw new ArgumentNullException(nameof(opt));
        }

        // ============================================================
        // Public generation entry point
        // ============================================================
        // ============================================================
// Public generation entry point
// ============================================================
public void Generate(string pptPath, string outputPdfPath, IProgress<string>? progress = null)
{
    ValidateInputs(pptPath, outputPdfPath);

    if (_opt.AlwaysUseTempLocalCopy)
    {
        progress?.Report("Using local temp copy mode.");
        GenerateCore(pptPath, outputPdfPath, useTempCopy: true, progress);
        return;
    }

    try
    {
        GenerateCore(pptPath, outputPdfPath, useTempCopy: false, progress);
    }
    catch (Exception ex) when (IsRetryablePowerPointComFailure(ex))
    {
        progress?.Report("Direct processing failed due to a PowerPoint COM/RPC issue.");
        progress?.Report("Retrying using a local temp copy...");

        GenerateCore(pptPath, outputPdfPath, useTempCopy: true, progress);
    }
}

private void GenerateCore(
    string pptPath,
    string outputPdfPath,
    bool useTempCopy,
    IProgress<string>? progress = null)
{
    PowerPoint.Application? pptApp = null;
    PowerPoint.Presentation? pres = null;

    bool powerPointWasAlreadyRunning = false;

    string tempWorkDir = Path.Combine(
        Path.GetTempPath(),
        "ppt_handout_maker",
        Guid.NewGuid().ToString("N")
    );

    string localPptPath = Path.Combine(
        tempWorkDir,
        Path.GetFileName(pptPath)
    );

    string presentationToOpen = pptPath;

    string tempImageDir = useTempCopy
        ? Path.Combine(tempWorkDir, "images")
        : Path.Combine(
            Path.GetTempPath(),
            "ppt_handout_maker",
            Guid.NewGuid().ToString("N")
        );

    try
    {
        if (useTempCopy)
        {
            progress?.Report("Preparing local temp copy...");
            Directory.CreateDirectory(tempWorkDir);

            progress?.Report("Copying presentation to local temp folder...");
            File.Copy(pptPath, localPptPath, overwrite: true);

            presentationToOpen = localPptPath;
        }

        Directory.CreateDirectory(tempImageDir);

        progress?.Report("Starting PowerPoint...");

        powerPointWasAlreadyRunning = PowerPointInterop.TryGetRunningPowerPoint(out pptApp);

        if (pptApp == null)
        {
            powerPointWasAlreadyRunning = false;
            pptApp = new PowerPoint.Application();
        }

        pptApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;

        progress?.Report(useTempCopy ? "Opening local temp copy..." : "Opening presentation...");
        pres = pptApp.Presentations.Open(
            FileName: presentationToOpen,
            ReadOnly: Office.MsoTriState.msoTrue,
            Untitled: Office.MsoTriState.msoFalse,
            WithWindow: Office.MsoTriState.msoFalse
        );

        int slideCount = pres.Slides.Count;
        progress?.Report($"Slide count: {slideCount}");

        float slideRatioHW = pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth;
        progress?.Report($"Slide ratio (H/W): {slideRatioHW:0.0000}");

        var items = new List<(int SlideNumber, string SlideImagePath, List<NoteLine> NotesLines)>();

        for (int i = 1; i <= slideCount; i++)
        {
            progress?.Report($"SLIDE_PROGRESS|{i}|{slideCount}");

            PowerPoint.Slide? currentSlide = null;
            try
            {
                currentSlide = pres.Slides[i];

                string imagePath = SlideExport.ExportSlidePng(
                    currentSlide,
                    tempImageDir,
                    widthPx: _opt.SlideExportWidthPx);

                var notesLines = NotesExtractor.GetSlideNotesLines(currentSlide);

                if (_opt.SkipSlidesWithNoNotes && notesLines.Count == 0)
                    continue;

                items.Add((i, imagePath, notesLines));
            }
            finally
            {
                PowerPointInterop.FinalRelease(currentSlide);
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

        progress?.Report("Cleaning temp files...");
        try
        {
            if (useTempCopy)
            {
                if (Directory.Exists(tempWorkDir))
                    Directory.Delete(tempWorkDir, recursive: true);
            }
            else
            {
                if (Directory.Exists(tempImageDir))
                    Directory.Delete(tempImageDir, recursive: true);
            }
        }
        catch { }

        progress?.Report("Cleanup complete.");
    }
}

private static bool IsRetryablePowerPointComFailure(Exception ex)
{
    if (ex is COMException)
        return true;

    if (ex is InvalidOperationException && ex.InnerException is COMException)
        return true;

    return false;
}

        // ============================================================
        // Input validation
        // ============================================================
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
        // PowerPoint session and COM cleanup
        // ============================================================
        private static class PowerPointInterop
        {
            [DllImport("oleaut32.dll", PreserveSig = false)]
            private static extern void GetActiveObject(
                ref Guid rclsid,
                IntPtr reserved,
                [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

            public static bool TryGetRunningPowerPoint(out PowerPoint.Application? app)
            {
                app = null;

                try
                {
                    Guid clsidPowerPoint = new("91493441-5A91-11CF-8700-00AA0060263B");
                    GetActiveObject(ref clsidPowerPoint, IntPtr.Zero, out object obj);
                    app = (PowerPoint.Application)obj;
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
        // Slide image export
        // ============================================================
        private static class SlideExport
        {
            public static string ExportSlidePng(Slide slide, string outDir, int widthPx = 1280)
            {
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
        // Notes extraction and list reconstruction
        // ============================================================
        private static class NotesExtractor
        {
            // ------------------------------------------------------------
            // Public entry point
            // ------------------------------------------------------------
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

                var result = new List<NoteLine>(raw.Count);
                foreach (var p in raw)
                {
                    int level = p.IsListItem ? Math.Max(1, p.IndentLevel) : 0;
                    result.Add(new NoteLine(level, p.Text));
                }

                return result;
            }

            // ------------------------------------------------------------
            // Notes page scanning and text extraction
            // ------------------------------------------------------------
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
                            RemoveCountersAbove(numberCountersByIndent, indentLevel);

                            int current;
                            if (!numberCountersByIndent.TryGetValue(indentLevel, out current))
                            {
                                current = GetBulletStartValueOrDefault(para, 1);
                            }
                            else
                            {
                                current += 1;
                            }

                            numberCountersByIndent[indentLevel] = current;

                            object? styleObj = null;
                            try { styleObj = para.ParagraphFormat.Bullet.Style; } catch { }

                            prefix = FormatNumberLabel(current, styleObj) + " ";
                        }
                        else if (isBulleted)
                        {
                            RemoveCountersAtOrAbove(numberCountersByIndent, indentLevel);

                            if (!StartsWithBullet(text))
                                prefix = "• ";
                        }
                        else
                        {
                            RemoveCountersAtOrAbove(numberCountersByIndent, indentLevel);
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

            // ------------------------------------------------------------
            // Numbering counter maintenance
            // ------------------------------------------------------------
            private static void RemoveCountersAbove(Dictionary<int, int> counters, int indentLevel)
            {
                if (counters.Count == 0)
                    return;

                var keysToRemove = new List<int>();

                foreach (int key in counters.Keys)
                {
                    if (key > indentLevel)
                        keysToRemove.Add(key);
                }

                foreach (int key in keysToRemove)
                    counters.Remove(key);
            }

            private static void RemoveCountersAtOrAbove(Dictionary<int, int> counters, int indentLevel)
            {
                if (counters.Count == 0)
                    return;

                var keysToRemove = new List<int>();

                foreach (int key in counters.Keys)
                {
                    if (key >= indentLevel)
                        keysToRemove.Add(key);
                }

                foreach (int key in keysToRemove)
                    counters.Remove(key);
            }

            // ------------------------------------------------------------
            // Paragraph classification
            // ------------------------------------------------------------
            private static bool StartsWithBullet(string s)
            {
                if (string.IsNullOrEmpty(s))
                    return false;

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

            // ------------------------------------------------------------
            // Number label formatting
            // ------------------------------------------------------------
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
                if (value < 1)
                    value = 1;

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
                if (value < 1)
                    value = 1;

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
        // PDF layout and rendering
        // ============================================================
        private static class PdfBuilder
        {
            private static bool _isInitialized;
            private static readonly object InitLock = new();

            private readonly record struct LayoutSettings(
                float ThumbBorder,
                float ThumbPadding,
                float CellBorder,
                float BaseThumbWidth,
                float CompactThumbWidthScale,
                float CompactHeightScale,
                float SlideNumberRowHeight,
                float IndentStep);

            private static void EnsureInitialized()
            {
                if (_isInitialized)
                    return;

                lock (InitLock)
                {
                    if (_isInitialized)
                        return;

                    FontManager.RegisterFontWithCustomName("Arial", File.OpenRead("C:\\Windows\\Fonts\\arial.ttf"));
                    FontManager.RegisterFontWithCustomName("Arial", File.OpenRead("C:\\Windows\\Fonts\\arialbd.ttf"));

                    QuestPDF.Settings.License = LicenseType.Community;

                    _isInitialized = true;
                }
            }

            public static void BuildHandoutPdf(
                HandoutOptions opt,
                List<(int SlideNumber, string SlideImagePath, List<NoteLine> NotesLines)> items,
                string outputPdfPath,
                float slideRatioHW)
            {
                // ------------------------------------------------------------
                // Initialization and normalized inputs
                // ------------------------------------------------------------
                EnsureInitialized();

                string className = (opt.ClassName ?? "").Trim();
                string pdfTitle = (opt.PdfTitle ?? "").Trim();

                // ------------------------------------------------------------
                // Shared layout constants
                // ------------------------------------------------------------
                var layout = new LayoutSettings(
                    ThumbBorder: 1f,
                    ThumbPadding: 2f,
                    CellBorder: 1f,
                    BaseThumbWidth: 240f,
                    CompactThumbWidthScale: 0.7f,
                    CompactHeightScale: 0.45f,
                    SlideNumberRowHeight: 14f,
                    IndentStep: 14f);

                // ------------------------------------------------------------
                // Document shell
                // ------------------------------------------------------------
                Document.Create(container =>
                {
                    container.Page(page =>
                    {
                        page.Size(PageSizes.Letter);
                        page.Margin(48);
                        page.DefaultTextStyle(x => x
                            .FontFamily("Arial")
                            .FontSize(9));

                        // ------------------------------------------------------------
                        // Repeating header and footer
                        // ------------------------------------------------------------
                        page.Header().ShowIf(ctx => ctx.PageNumber > 1).Element(h =>
                        {
                            h.PaddingBottom(8)
                             .AlignRight()
                             .Text(className)
                             .FontSize(9)
                             .Bold();
                        });

                        page.Footer().ShowIf(ctx => ctx.PageNumber > 1).Element(f =>
                        {
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

                        // ------------------------------------------------------------
                        // Main page content
                        // ------------------------------------------------------------
                        page.Content()
                            .PaddingTop(12)
                            .Column(col =>
                            {
                                // First-page title block
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

                                // Render one handout block per slide
                                for (int i = 0; i < items.Count; i++)
                                {
                                    var (slideNumber, imgPath, notesLines) = items[i];

                                    ComposeSlideBlock(
                                        col,
                                        slideNumber,
                                        imgPath,
                                        notesLines,
                                        slideRatioHW,
                                        showDividerAfter: i < items.Count - 1,
                                        layout,
                                        opt);
                                }
                            });
                    });
                })
                .GeneratePdf(outputPdfPath);
            }

            private static void ComposeSlideBlock(
                ColumnDescriptor col,
                int slideNumber,
                string imgPath,
                List<NoteLine> notesLines,
                float slideRatioHW,
                bool showDividerAfter,
                LayoutSettings layout,
                HandoutOptions opt)
            {
                // Read item state
                bool hasNotes = notesLines.Count > 0;

                // Compute slide block dimensions
                float effectiveThumbWidth = hasNotes
                    ? layout.BaseThumbWidth
                    : layout.BaseThumbWidth * layout.CompactThumbWidthScale;

                float imageHeight = effectiveThumbWidth * slideRatioHW;

                float thumbBoxHeight =
                    imageHeight
                    + (2 * layout.ThumbPadding)
                    + (2 * layout.ThumbBorder)
                    - layout.CellBorder;

                float compactThumbBoxHeight = thumbBoxHeight * layout.CompactHeightScale;
                float effectiveHeight = hasNotes ? thumbBoxHeight : compactThumbBoxHeight;

                // Render slide thumbnail and notes columns
                col.Item()
                   .ShowEntire()
                   .PaddingVertical(hasNotes ? 20 : 6)
                   .Element(outer =>
                   {
                       var cell = outer
                           .Border(0)
                           .BorderColor(Colors.Grey.Darken1)
                           .Padding(0);

                       cell = hasNotes
                           ? cell.MinHeight(effectiveHeight)
                           : cell.Height(effectiveHeight);

                       cell.Element(block =>
                       {
                           block.Row(row =>
                           {
                               // Left column: slide label and thumbnail
                               row.RelativeItem(1.0f).Element(left =>
                               {
                                   float thumbAreaHeight = Math.Max(1f, effectiveHeight - layout.SlideNumberRowHeight);

                                   left.Height(effectiveHeight)
                                       .AlignTop()
                                       .Column(lc =>
                                       {
                                           lc.Spacing(0);

                                           lc.Item()
                                             .Height(layout.SlideNumberRowHeight)
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

                               // Right column: instructor notes
                               row.RelativeItem(1.2f).Element(right =>
                               {
                                   right.Column(ncol =>
                                   {
                                       ComposeNotesColumn(ncol, notesLines, layout.IndentStep, opt);
                                   });
                               });
                           });
                       });
                   });

                // Divider between slide sections
                if (showDividerAfter)
                {
                    col.Item()
                       .LineHorizontal(1)
                       .LineColor(Colors.Grey.Lighten2);
                }
            }

            private static void ComposeNotesColumn(
                ColumnDescriptor ncol,
                List<NoteLine> notesLines,
                float indentStep,
                HandoutOptions opt)
            {
                ncol.Spacing(2);

                // Empty-notes case
                if (notesLines.Count == 0)
                {
                    if (opt.ShowNoNotesPlaceholder)
                        ncol.Item().Text("(No notes)");

                    return;
                }

                // Notes layout metrics
                float gutterWidth = ComputeGutterWidthForSlide(notesLines);

                // Note line rendering loop
                foreach (var nl in notesLines)
                {
                    string line = string.IsNullOrWhiteSpace(nl.Text) ? " " : nl.Text;

                    int indentLevels = Math.Max(0, nl.Level - 1);
                    float leftPad = indentLevels * indentStep;

                    // Plain paragraph
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

                    // List item with aligned prefix/body columns
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
            }

            private static float ComputeGutterWidthForSlide(List<NoteLine> notesLines)
            {
                const float baseWidth = 5f;
                const float perChar = 4.0f;
                const float minWidth = 12f;
                const float maxWidth = 32f;

                if (notesLines.Count == 0)
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
                        firstToken.EndsWith('.') ||
                        firstToken.EndsWith(')') ||
                        (firstToken.StartsWith('(') && firstToken.EndsWith(')'));

                    if (looksNumbered)
                        return (firstToken, rest);
                }

                return ("", s);
            }
        }

        // ============================================================
        // Legacy helpers currently unused
        // These helpers are retained for reference and are not part
        // of the active handout generation path.
        // ============================================================

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
            boxHeight -= cellBorderPts;

            return boxHeight;
        }

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

                int v = Convert.ToInt32(visible);
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