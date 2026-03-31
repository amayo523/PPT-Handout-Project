using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using PptNotesHandoutMaker.Core;

namespace PptNotesHandoutMaker
{
    public partial class MainWindow : Window
    {
        private const bool ENABLE_SHARED_DRIVE_WARNING = false;
        private bool _isGenerating;

        public MainWindow()
        {
            InitializeComponent();
            UpdateUiState();
        }

        // -----------------------------
        // Browse PowerPoint button
        // -----------------------------
        private async void BrowsePpt_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Select PowerPoint file",
                Filter = "PowerPoint files (*.pptx;*.ppt)|*.pptx;*.ppt|All files (*.*)|*.*"
            };

            if (dlg.ShowDialog() != true)
                return;

            string selectedPath = dlg.FileName;

            if (ENABLE_SHARED_DRIVE_WARNING && !ConfirmSharedDriveWarning(this, selectedPath))
                return;

            PptPathBox.Text = selectedPath;

            OutPathBox.Text = Path.Combine(
                Path.GetDirectoryName(selectedPath)!,
                Path.GetFileNameWithoutExtension(selectedPath) + "_Instructor Guide.pdf"
            );

            AppendStatus("Selected PowerPoint:");
            AppendStatus(selectedPath);

            AppendStatus("Reading title from first slide...");

            try
            {
                string? detectedTitle = await StaTask.Run(() =>
                    PowerPointTitleReader.TryReadFirstSlideTitle(selectedPath));

                if (!string.IsNullOrWhiteSpace(detectedTitle))
                {
                    PdfTitleBox.Text = detectedTitle;

                    AppendStatus("Detected PDF title:");
                    AppendStatus(detectedTitle);
                }
                else
                {
                    AppendStatus("No usable title found on slide 1.");
                }
            }
            catch (Exception ex)
            {
                AppendStatus("Error reading slide 1 title:");
                AppendStatus(ex.Message);
            }

            UpdateUiState();
        }

        // -----------------------------
        // Browse Output PDF button
        // -----------------------------
        private void BrowseOut_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Choose output PDF",
                Filter = "PDF files (*.pdf)|*.pdf",
                DefaultExt = "pdf",
                AddExtension = true,
                FileName = string.IsNullOrWhiteSpace(OutPathBox.Text)
                    ? "handout.pdf"
                    : Path.GetFileName(OutPathBox.Text)
            };

            // Nice UX: start in PPT folder if available
            if (!string.IsNullOrWhiteSpace(PptPathBox.Text))
            {
                try { dlg.InitialDirectory = Path.GetDirectoryName(PptPathBox.Text); } catch { }
            }

            if (dlg.ShowDialog() != true)
                return;

            OutPathBox.Text = dlg.FileName;

            AppendStatus("Output PDF:");
            AppendStatus(dlg.FileName);

            UpdateUiState();
        }

        // -----------------------------
        // Generate button
        // -----------------------------
        private async void Generate_Click(object sender, RoutedEventArgs e)
        {
            if (_isGenerating)
                return;

            string className = (ClassNameBox.Text ?? "").Trim();
            string pdfTitle = (PdfTitleBox.Text ?? "").Trim();

            if (string.IsNullOrWhiteSpace(className) || string.IsNullOrWhiteSpace(pdfTitle))
            {
                MessageBox.Show(
                    this,
                    "Please enter both a Class Name and a PDF Title before generating the PDF.",
                    "Missing Course Information",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );

                return;
            }

            string pptPath = (PptPathBox.Text ?? "").Trim();
            string outPath = (OutPathBox.Text ?? "").Trim();

            // Final validation (even though button should be disabled if invalid)
            if (!File.Exists(pptPath))
            {
                AppendStatus("PowerPoint file not found.");
                UpdateUiState();
                return;
            }

            if (string.IsNullOrWhiteSpace(outPath))
            {
                AppendStatus("Choose an output PDF path.");
                UpdateUiState();
                return;
            }

            // Ensure we don't overwrite existing
            outPath = GetNextAvailableFilePath(outPath);
            OutPathBox.Text = outPath;

            StatusBox.Clear();

            _isGenerating = true;
            UpdateUiState(); // disables Generate, keeps Browse clickable

            IProgress<string> progress = new Progress<string>(AppendStatus);

            try
            {
                progress.Report("Starting generation...");

                var opt = new HandoutOptions
                {
                    ClassName = (ClassNameBox.Text ?? "").Trim(),
                    PdfTitle = (PdfTitleBox.Text ?? "").Trim()
                };

                var gen = new HandoutGenerator(opt);

                await StaTask.Run(() =>
                {
                    gen.Generate(pptPath, outPath, progress);
                });

                // Open the output PDF
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outPath)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                progress.Report("ERROR:");
                progress.Report(ex.ToString());
            }
            finally
            {
                _isGenerating = false;
                UpdateUiState();
            }
        }

        // -----------------------------
        // Centralized UI enable/disable logic
        // -----------------------------
        private void UpdateUiState()
        {
            bool hasValidPpt = !string.IsNullOrWhiteSpace(PptPathBox.Text) && File.Exists(PptPathBox.Text);
            bool hasOut = !string.IsNullOrWhiteSpace(OutPathBox.Text);
            bool hasClass = !string.IsNullOrWhiteSpace(ClassNameBox.Text);
            bool hasPdfTitle = !string.IsNullOrWhiteSpace(PdfTitleBox.Text);

            // Browse buttons remain clickable at all times (we never disable them).
            GenerateBtn.IsEnabled = !_isGenerating && hasValidPpt && hasOut && hasClass && hasPdfTitle;
        }

        // -----------------------------
        // Status helper
        // -----------------------------
        private void AppendStatus(string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => AppendStatus(message)));
                return;
            }

            StatusBox.AppendText(message + Environment.NewLine);
            StatusBox.ScrollToEnd();
            StatusBox.InvalidateVisual();
            Dispatcher.BeginInvoke(new Action(() => { }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void AnyInputChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateUiState();
        }

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

        static bool IsSharedDrivePath(string filePath, out string driveLabel)
        {
            driveLabel = "";

            if (string.IsNullOrWhiteSpace(filePath))
                return false;

            // UNC path (\\server\share\...)
            if (filePath.StartsWith(@"\\"))
            {
                driveLabel = TryGetUncShareRoot(filePath) ?? "network share";
                return true;
            }

            try
            {
                string root = Path.GetPathRoot(filePath) ?? "";
                if (string.IsNullOrWhiteSpace(root))
                    return false;

                var di = new DriveInfo(root);
                if (di.DriveType == DriveType.Network)
                {
                    driveLabel = root.TrimEnd('\\');
                    return true;
                }
            }
            catch { }

            return false;
        }

        static string? TryGetUncShareRoot(string uncPath)
        {
            try
            {
                var parts = uncPath.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2)
                    return $@"\\{parts[0]}\{parts[1]}";
            }
            catch { }
            return null;
        }

        static bool ConfirmSharedDriveWarning(Window owner, string filePath)
        {
            if (!IsSharedDrivePath(filePath, out string driveName))
                return true;

            string msg =
                $"This file is on a shared drive ({driveName}).\n\n" +
                "The PDF may take longer to generate, or may fail.\n\n" +
                "Suggestion: Save a local copy (e.g., Desktop/Documents) and use that instead.\n\n" +
                "Are you sure you wish to continue?";

            var result = MessageBox.Show(
                owner,
                msg,
                "Shared Drive Warning",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            return result == MessageBoxResult.Yes;
        }
    }
}
