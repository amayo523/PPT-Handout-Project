using System.Collections.Generic;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;

namespace PptNotesHandoutMaker.Core
{
    internal static class PowerPointTitleReader
    {
        public static string? TryReadFirstSlideTitle(string pptPath)
        {
            if (string.IsNullOrWhiteSpace(pptPath) || !File.Exists(pptPath))
                return null;

            PowerPoint.Application? pptApp = null;
            PowerPoint.Presentation? pres = null;
            PowerPoint.Slide? slide = null;

            bool powerPointWasAlreadyRunning = false;

            try
            {
                powerPointWasAlreadyRunning = PowerPointInteropUtil.TryGetRunningPowerPoint(out pptApp);

                if (pptApp == null)
                {
                    powerPointWasAlreadyRunning = false;
                    pptApp = new PowerPoint.Application();
                }

                pptApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;

                pres = pptApp.Presentations.Open(
                    FileName: pptPath,
                    ReadOnly: Office.MsoTriState.msoTrue,
                    Untitled: Office.MsoTriState.msoFalse,
                    WithWindow: Office.MsoTriState.msoFalse
                );

                if (pres.Slides.Count < 1)
                    return null;

                slide = pres.Slides[1];
                return ExtractTitleIgnoringLeftQuarter(slide);
            }
            finally
            {
                if (slide != null)
                    PowerPointInteropUtil.FinalRelease(slide);

                if (pres != null)
                {
                    try { pres.Close(); } catch { }
                    PowerPointInteropUtil.FinalRelease(pres);
                }

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
                        PowerPointInteropUtil.FinalRelease(pptApp);
                    }
                }
            }
        }

        private static string NormalizeTitle(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            text = text.Trim();

            return Regex.Replace(text, @"^\s*Module\b\s*", "Mod ", RegexOptions.IgnoreCase);
        }

        private static string? ExtractTitleIgnoringLeftQuarter(PowerPoint.Slide slide)
        {
            float slideWidth = 0f;

            try
            {
                var pres = slide.Parent as PowerPoint.Presentation;
                slideWidth = pres?.PageSetup.SlideWidth ?? 0f;
            }
            catch { }

            float cutoff = slideWidth * 0.25f;
            var fallbackTexts = new List<string>();

            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                PowerPoint.Shape? sh = null;

                try
                {
                    sh = slide.Shapes[i];

                    if (sh.HasTextFrame != Office.MsoTriState.msoTrue)
                        continue;

                    var tf = sh.TextFrame;
                    if (tf == null || tf.HasText != Office.MsoTriState.msoTrue)
                        continue;

                    string text = (tf.TextRange?.Text ?? "")
                        .Replace("\r", " ")
                        .Replace("\n", " ")
                        .Trim();

                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    float shapeCenterX = sh.Left + (sh.Width / 2f);

                    if (slideWidth > 0f && shapeCenterX < cutoff)
                        continue;

                    if (IsTitlePlaceholder(sh))
                        return NormalizeTitle(text);

                    fallbackTexts.Add(text);
                }
                finally
                {
                    if (sh != null)
                        PowerPointInteropUtil.FinalRelease(sh);
                }
            }

            return fallbackTexts.Count > 0 ? NormalizeTitle(fallbackTexts[0]) : null;
        }

        private static bool IsTitlePlaceholder(PowerPoint.Shape sh)
        {
            try
            {
                if (sh.Type != Office.MsoShapeType.msoPlaceholder)
                    return false;

                var phType = sh.PlaceholderFormat.Type;
                return phType == PowerPoint.PpPlaceholderType.ppPlaceholderTitle
                    || phType == PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle;
            }
            catch
            {
                return false;
            }
        }
    }
}