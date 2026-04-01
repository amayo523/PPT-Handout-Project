using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Versioning;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using PptNotesHandoutMaker.Core;

namespace PptNotesHandoutMaker
{
    public partial class MainWindow : Window
    {
        private static readonly bool FORCE_FILENAME_FALLBACK_FOR_TEST = false;
        private const bool ENABLE_SHARED_DRIVE_WARNING = false;

        private bool _isGenerating;
        private bool _isReadingTitles;

        private readonly List<BatchPptItem> _selectedPpts = new();

        public MainWindow()
        {
            InitializeComponent();
            UpdatePptCountDisplay();
            UpdateDropHintVisibility();
            UpdateDropZoneVisual(DropZoneState.Normal);
            UpdateUiState();
        }

        private void AnyOptionChanged(object sender, RoutedEventArgs e)
        {
            UpdateUiState();
        }

        // -----------------------------
        // Add PPT button
        // -----------------------------
        private async void BrowsePpt_Click(object sender, RoutedEventArgs e)
        {
            if (_isGenerating || _isReadingTitles)
                return;

            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select PowerPoint files",
                Filter = "PowerPoint files (*.pptx;*.ppt)|*.pptx;*.ppt|All files (*.*)|*.*",
                Multiselect = true
            };

            if (dlg.ShowDialog() != true || dlg.FileNames.Length == 0)
                return;

            await AddPowerPointsAsync(dlg.FileNames);
        }

        // -----------------------------
        // Drag and drop support
        // -----------------------------
        private void Window_PreviewDragOver(object sender, System.Windows.DragEventArgs e)
        {
            if (_isGenerating || _isReadingTitles)
            {
                UpdateDropZoneVisual(DropZoneState.Normal);
                e.Effects = System.Windows.DragDropEffects.None;
                e.Handled = true;
                return;
            }

            if (!e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                UpdateDropZoneVisual(DropZoneState.Invalid);
                e.Effects = System.Windows.DragDropEffects.None;
                e.Handled = true;
                return;
            }

            var paths = e.Data.GetData(System.Windows.DataFormats.FileDrop) as string[];
            if (paths == null || paths.Length == 0)
            {
                UpdateDropZoneVisual(DropZoneState.Invalid);
                e.Effects = System.Windows.DragDropEffects.None;
                e.Handled = true;
                return;
            }

            bool allArePowerPoints = paths.All(IsPowerPointFile);

            UpdateDropZoneVisual(allArePowerPoints ? DropZoneState.Valid : DropZoneState.Invalid);
            e.Effects = allArePowerPoints
                ? System.Windows.DragDropEffects.Copy
                : System.Windows.DragDropEffects.None;

            e.Handled = true;
        }

        private async void Window_Drop(object sender, System.Windows.DragEventArgs e)
        {
            try
            {
                if (_isGenerating || _isReadingTitles)
                    return;

                if (!e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
                    return;

                var droppedPaths = e.Data.GetData(System.Windows.DataFormats.FileDrop) as string[];
                if (droppedPaths == null || droppedPaths.Length == 0)
                    return;

                // At this point, PreviewDragOver already guaranteed all files are valid
                await AddPowerPointsAsync(droppedPaths);
            }
            finally
            {
                UpdateDropZoneVisual(DropZoneState.Normal);
            }
        }

        private void Window_DragLeave(object sender, System.Windows.DragEventArgs e)
        {
            UpdateDropZoneVisual(DropZoneState.Normal);
        }

        private static bool IsPowerPointFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            string ext = Path.GetExtension(path);
            return ext.Equals(".pptx", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".ppt", StringComparison.OrdinalIgnoreCase);
        }

        private void UpdateDropZoneVisual(DropZoneState state)
        {
            switch (state)
            {
                case DropZoneState.Normal:
                    DropZoneBorder.BorderBrush = System.Windows.Media.Brushes.LightGray;
                    DropZoneBorder.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(250, 250, 250));
                    break;

                case DropZoneState.Valid:
                    DropZoneBorder.BorderBrush = System.Windows.Media.Brushes.ForestGreen;
                    DropZoneBorder.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(240, 255, 240));
                    break;

                case DropZoneState.Invalid:
                    DropZoneBorder.BorderBrush = System.Windows.Media.Brushes.IndianRed;
                    DropZoneBorder.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 240, 240));
                    break;
            }
        }

        private enum DropZoneState
        {
            Normal,
            Valid,
            Invalid
        }

        // -----------------------------
        // Shared add/read-title flow
        // -----------------------------
        private async Task AddPowerPointsAsync(string[] filePaths)
        {
            if (filePaths == null || filePaths.Length == 0)
                return;

            // Default output folder = folder of the last selected PowerPoint
            string lastSelectedPath = filePaths[filePaths.Length - 1];
            string? lastSelectedFolder = Path.GetDirectoryName(lastSelectedPath);

            if (!string.IsNullOrWhiteSpace(lastSelectedFolder) && Directory.Exists(lastSelectedFolder))
            {
                OutPathBox.Text = lastSelectedFolder;
            }

            _isReadingTitles = true;
            UpdateUiState();
            UpdateDropHintVisibility();

            try
            {
                AppendStatus("--------------------------------------------------");
                StatusLabel.Text = $"Status: Starting title read (0 of {filePaths.Length})";

                for (int i = 0; i < filePaths.Length; i++)
                {
                    string selectedPath = filePaths[i];

                    if (ENABLE_SHARED_DRIVE_WARNING && !ConfirmSharedDriveWarning(this, selectedPath))
                        continue;

                    bool alreadyAdded = _selectedPpts.Any(x =>
                        string.Equals(x.PptPath, selectedPath, StringComparison.OrdinalIgnoreCase));

                    if (alreadyAdded)
                    {
                        AppendStatus($"Skipped already-added file: {selectedPath}");
                        continue;
                    }

                    string displayName = Path.GetFileName(selectedPath);
                    if (displayName.Length > 60)
                        displayName = displayName.Substring(0, 57) + "...";

                    int currentIndex = i + 1;
                    int totalCount = filePaths.Length;

                    StatusLabel.Text = $"Status: Reading titles ({currentIndex} of {totalCount}) - {displayName}";

                    string detectedTitle = "";
                    bool usedFallback = false;

                    if (FORCE_FILENAME_FALLBACK_FOR_TEST)
                    {
                        detectedTitle = Path.GetFileNameWithoutExtension(selectedPath);
                        usedFallback = true;

                        AppendStatus($"TEST MODE: forcing filename fallback for {selectedPath}");
                    }
                    else
                    {
                        try
                        {
                            detectedTitle = await StaTask.Run(() =>
                                PowerPointTitleReader.TryReadFirstSlideTitle(selectedPath)) ?? "";
                        }
                        catch (Exception ex)
                        {
                            AppendStatus($"Error reading title for {selectedPath}");
                            AppendStatus(ex.Message);

                            detectedTitle = Path.GetFileNameWithoutExtension(selectedPath);
                            usedFallback = true;

                            AppendStatus($"Using filename as PDF title: {detectedTitle}");
                        }

                        if (string.IsNullOrWhiteSpace(detectedTitle))
                        {
                            detectedTitle = Path.GetFileNameWithoutExtension(selectedPath);
                            usedFallback = true;

                            AppendStatus($"No slide title found. Using filename as PDF title: {detectedTitle}");
                        }
                    }

                    _selectedPpts.Add(new BatchPptItem
                    {
                        PptPath = selectedPath,
                        PdfTitle = detectedTitle,
                        UsedFilenameFallback = usedFallback
                    });

                    AppendStatus($"Added: {selectedPath}");
                }

                if (_selectedPpts.Count == 0)
                {
                    UpdatePptCountDisplay();
                    UpdateDropHintVisibility();
                    UpdateUiState();
                    StatusLabel.Text = $"Status: Added {_selectedPpts.Count} PowerPoint file(s)";
                    return;
                }

                RebuildBatchTitlesUi();
                UpdatePptCountDisplay();
                UpdateDropHintVisibility();
            }
            finally
            {
                _isReadingTitles = false;
                UpdateUiState();
            }
        }

        private void UpdatePptCountDisplay()
        {
            int count = _selectedPpts.Count;

            if (count > 0)
            {
                PptCountText.Text = count == 1
                    ? "1 file selected"
                    : $"{count} files selected";

                PptCountText.Visibility = Visibility.Visible;
            }
            else
            {
                PptCountText.Text = "";
                PptCountText.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdateDropHintVisibility()
        {
            DropHintText.Visibility = _selectedPpts.Count == 0
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        // -----------------------------
        // Browse Output PDF button
        // -----------------------------
        [SupportedOSPlatform("windows")]
        private void BrowseOut_Click(object sender, RoutedEventArgs e)
        {
            using var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Choose destination folder for generated PDFs",
                UseDescriptionForTitle = true
            };

            if (_selectedPpts.Count > 0)
            {
                try
                {
                    string firstDir = Path.GetDirectoryName(_selectedPpts[0].PptPath) ?? "";
                    if (Directory.Exists(firstDir))
                        dlg.InitialDirectory = firstDir;
                }
                catch { }
            }

            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK || string.IsNullOrWhiteSpace(dlg.SelectedPath))
                return;

            OutPathBox.Text = dlg.SelectedPath;

            AppendStatus("Output folder:");
            AppendStatus(dlg.SelectedPath);

            UpdateUiState();
        }

        // -----------------------------
        // Generate button
        // -----------------------------
        private async void Generate_Click(object sender, RoutedEventArgs e)
        {
            if (_isGenerating || _isReadingTitles)
                return;

            string className = (ClassNameBox.Text ?? "").Trim();
            string outputFolder = (OutPathBox.Text ?? "").Trim();
            bool alwaysUseTempLocalCopy = AlwaysUseTempCopyCheckBox.IsChecked == true;

            if (string.IsNullOrWhiteSpace(className))
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Please enter a Class Name before generating PDFs.",
                    "Missing Course Information",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                return;
            }

            if (_selectedPpts.Count == 0)
            {
                AppendStatus("No PowerPoint files selected.");
                UpdateUiState();
                return;
            }

            if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
            {
                AppendStatus("Choose a valid destination folder.");
                UpdateUiState();
                return;
            }

            StatusBox.Clear();

            _isGenerating = true;
            UpdateUiState();

            IProgress<string> progress = new Progress<string>(msg =>
            {
                if (msg.StartsWith("SLIDE_PROGRESS|"))
                {
                    var parts = msg.Split('|');
                    int current = int.Parse(parts[1]);
                    int total = int.Parse(parts[2]);

                    ReplaceLastStatusLine($"Exporting slide {current}/{total}...");
                }
                else
                {
                    AppendStatus(msg);
                }
            });

            try
            {
                progress.Report($"Starting batch generation for {_selectedPpts.Count} file(s)...");
                StatusLabel.Text = $"Status: Starting batch (0 of {_selectedPpts.Count})";

                await StaTask.Run(() =>
                {
                    for (int i = 0; i < _selectedPpts.Count; i++)
                    {
                        var item = _selectedPpts[i];

                        string pptPath = (item.PptPath ?? "").Trim();
                        string pdfTitle = (item.PdfTitle ?? "").Trim();

                        string displayName = Path.GetFileName(pptPath);
                        if (displayName.Length > 60)
                            displayName = displayName.Substring(0, 57) + "...";

                        int currentIndex = i + 1;
                        int totalCount = _selectedPpts.Count;

                        Dispatcher.Invoke(() =>
                        {
                            StatusLabel.Text = $"Status: Processing ({currentIndex} of {totalCount}) - {displayName}";
                        });

                        try
                        {
                            progress.Report("--------------------------------------------------");
                            progress.Report($"Processing: {Path.GetFileName(pptPath)}");

                            if (string.IsNullOrWhiteSpace(pptPath) || !File.Exists(pptPath))
                            {
                                progress.Report("Skipped: file not found.");
                                continue;
                            }

                            if (string.IsNullOrWhiteSpace(pdfTitle))
                                pdfTitle = Path.GetFileNameWithoutExtension(pptPath);

                            string outputPdfPath = Path.Combine(
                                outputFolder,
                                $"{Path.GetFileNameWithoutExtension(pptPath)}_Instructor Guide.pdf"
                            );

                            outputPdfPath = GetNextAvailableFilePath(outputPdfPath);

                            progress.Report($"PDF title: {pdfTitle}");
                            progress.Report($"Output: {outputPdfPath}");

                            var opt = new HandoutOptions
                            {
                                ClassName = className,
                                PdfTitle = pdfTitle,
                                ShowNoNotesPlaceholder = true,
                                AlwaysUseTempLocalCopy = alwaysUseTempLocalCopy
                            };

                            var gen = new HandoutGenerator(opt);
                            gen.Generate(pptPath, outputPdfPath, progress);

                            progress.Report($"Finished: {outputPdfPath}");
                        }
                        catch (Exception ex)
                        {
                            progress.Report("ERROR processing file:");
                            progress.Report(pptPath);
                            progress.Report(ex.ToString());
                        }
                    }
                });

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputFolder)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                progress.Report("BATCH ERROR:");
                progress.Report(ex.ToString());
            }
            finally
            {
                _isGenerating = false;
                StatusLabel.Text = "Status: Batch complete";
                UpdateUiState();
            }
        }

        // -----------------------------
        // Centralized UI enable/disable logic
        // -----------------------------
        private void UpdateUiState()
        {
            bool isBusy = _isGenerating || _isReadingTitles;

            bool hasValidPpts = _selectedPpts.Count > 0 && _selectedPpts.All(x => File.Exists(x.PptPath));
            bool hasOutFolder = !string.IsNullOrWhiteSpace(OutPathBox.Text) && Directory.Exists(OutPathBox.Text);
            bool hasClass = !string.IsNullOrWhiteSpace(ClassNameBox.Text);
            bool hasAllTitles = _selectedPpts.Count > 0 && _selectedPpts.All(x => !string.IsNullOrWhiteSpace(x.PdfTitle));

            GenerateBtn.IsEnabled = !isBusy && hasValidPpts && hasOutFolder && hasClass && hasAllTitles;

            AddPptBtn.IsEnabled = !isBusy;
            ClearBatchBtn.IsEnabled = !isBusy;
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

        // -----------------------------
        // Replaces the last status line for slide exporting
        // -----------------------------
        private void ReplaceLastStatusLine(string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => ReplaceLastStatusLine(message)));
                return;
            }

            string currentText = StatusBox.Text;

            if (string.IsNullOrWhiteSpace(currentText))
            {
                StatusBox.AppendText(message + Environment.NewLine);
            }
            else
            {
                var lines = currentText.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

                if (lines.Length > 1)
                    lines[lines.Length - 2] = message;
                else
                    lines[0] = message;

                StatusBox.Text = string.Join(Environment.NewLine, lines);
            }

            StatusBox.ScrollToEnd();
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

        private static bool IsSharedDrivePath(string filePath, out string driveLabel)
        {
            driveLabel = "";

            if (string.IsNullOrWhiteSpace(filePath))
                return false;

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

        private static string? TryGetUncShareRoot(string uncPath)
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

        private static bool ConfirmSharedDriveWarning(Window owner, string filePath)
        {
            if (!IsSharedDrivePath(filePath, out string driveName))
                return true;

            string msg =
                $"This file is on a shared drive ({driveName}).\n\n" +
                "The PDF may take longer to generate, or may fail.\n\n" +
                "Suggestion: Save a local copy (e.g., Desktop/Documents) and use that instead.\n\n" +
                "Are you sure you wish to continue?";

            var result = System.Windows.MessageBox.Show(
                owner,
                msg,
                "Shared Drive Warning",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            return result == MessageBoxResult.Yes;
        }

        // -----------------------------
        // Rebuilding Batch Titles in the UI
        // -----------------------------
        private void RebuildBatchTitlesUi()
        {
            BatchTitlesPanel.Children.Clear();

            foreach (var item in _selectedPpts.ToList())
            {
                var outerBorder = new System.Windows.Controls.Border
                {
                    BorderThickness = new Thickness(1),
                    Margin = new Thickness(0, 0, 0, 10),
                    Padding = new Thickness(8),
                    CornerRadius = new CornerRadius(4),
                    BorderBrush = item.UsedFilenameFallback
                        ? System.Windows.Media.Brushes.DarkOrange
                        : System.Windows.Media.Brushes.LightGray
                };

                var outer = new System.Windows.Controls.StackPanel();

                var topRow = new System.Windows.Controls.Grid();
                topRow.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition
                {
                    Width = new GridLength(1, GridUnitType.Star)
                });
                topRow.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition
                {
                    Width = GridLength.Auto
                });

                var fileBlock = new System.Windows.Controls.TextBlock
                {
                    Text = Path.GetFileName(item.PptPath),
                    ToolTip = item.PptPath,
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(0, 0, 8, 2),
                    VerticalAlignment = VerticalAlignment.Center
                };

                var removeBtn = new System.Windows.Controls.Button
                {
                    Content = "X",
                    Width = 26,
                    Height = 26,
                    Tag = item,
                    ToolTip = "Remove this PowerPoint from the batch"
                };
                removeBtn.Click += RemoveBatchItem_Click;

                System.Windows.Controls.Grid.SetColumn(fileBlock, 0);
                System.Windows.Controls.Grid.SetColumn(removeBtn, 1);

                topRow.Children.Add(fileBlock);
                topRow.Children.Add(removeBtn);

                var titleLabel = new System.Windows.Controls.TextBlock
                {
                    Margin = new Thickness(0, 2, 0, 2),
                    FontWeight = FontWeights.SemiBold
                };

                titleLabel.Inlines.Add(new Run("PDF Title:"));

                if (item.UsedFilenameFallback)
                {
                    titleLabel.Inlines.Add(new Run(" Used file name as fallback - review title ")
                    {
                        Foreground = System.Windows.Media.Brushes.DarkOrange,
                    });

                    var retryLink = new Hyperlink(new Run("[Retry]"))
                    {
                        Foreground = System.Windows.Media.Brushes.DarkOrange,
                        Cursor = System.Windows.Input.Cursors.Hand,
                        Tag = item,
                        TextDecorations = TextDecorations.Underline
                    };

                    retryLink.Click += RetryTitleRead_Link_Click;
                    titleLabel.Inlines.Add(retryLink);
                }

                var titleBox = new System.Windows.Controls.TextBox
                {
                    Text = item.PdfTitle,
                    Tag = item,
                    Margin = new Thickness(0, 0, 0, 0)
                };

                titleBox.TextChanged += BatchTitleBox_TextChanged;

                outer.Children.Add(topRow);
                outer.Children.Add(titleLabel);
                outer.Children.Add(titleBox);

                outerBorder.Child = outer;
                BatchTitlesPanel.Children.Add(outerBorder);
            }
        }

        private void RemoveBatchItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn &&
                btn.Tag is BatchPptItem item)
            {
                _selectedPpts.Remove(item);
                RebuildBatchTitlesUi();

                AppendStatus($"Removed: {Path.GetFileName(item.PptPath)}");
                UpdatePptCountDisplay();
                UpdateDropHintVisibility();
                UpdateUiState();
            }
        }

        private void BatchTitleBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (sender is System.Windows.Controls.TextBox tb &&
                tb.Tag is BatchPptItem item)
            {
                item.PdfTitle = tb.Text ?? "";
            }

            UpdateUiState();
        }

        // -----------------------------
        // Clear Button
        // -----------------------------
        private void ClearBatch_Click(object sender, RoutedEventArgs e)
        {
            _selectedPpts.Clear();
            BatchTitlesPanel.Children.Clear();

            StatusBox.Clear();
            AppendStatus("Batch list cleared.");

            UpdatePptCountDisplay();
            UpdateDropHintVisibility();
            UpdateUiState();
        }

        // -----------------------------
        // Retry Button
        // -----------------------------
        private async void RetryTitleRead_Link_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Hyperlink link || link.Tag is not BatchPptItem item)
                return;

            string pptPath = (item.PptPath ?? "").Trim();
            if (string.IsNullOrWhiteSpace(pptPath) || !File.Exists(pptPath))
            {
                AppendStatus("Retry failed: file not found.");
                AppendStatus(pptPath);
                return;
            }

            string displayName = Path.GetFileName(pptPath);
            if (displayName.Length > 60)
                displayName = displayName.Substring(0, 57) + "...";

            try
            {
                link.IsEnabled = false;

                StatusLabel.Text = $"Status: Retrying title read for {displayName}";

                AppendStatus("--------------------------------------------------");
                AppendStatus($"Retrying title read for: {pptPath}");

                string? retriedTitle = await StaTask.Run(() =>
                    PowerPointTitleReader.TryReadFirstSlideTitle(pptPath));

                if (!string.IsNullOrWhiteSpace(retriedTitle))
                {
                    item.PdfTitle = retriedTitle.Trim();
                    item.UsedFilenameFallback = false;

                    AppendStatus($"Retry succeeded. Updated PDF title to: {item.PdfTitle}");
                    StatusLabel.Text = $"Status: Retry succeeded for {displayName}";

                    RebuildBatchTitlesUi();
                    UpdateUiState();
                }
                else
                {
                    AppendStatus("Retry did not find a usable slide title.");
                    StatusLabel.Text = $"Status: Retry found no title for {displayName}";
                }
            }
            catch (Exception ex)
            {
                AppendStatus($"Retry failed for: {pptPath}");
                AppendStatus(ex.Message);
                StatusLabel.Text = $"Status: Retry failed for {displayName}";
            }
            finally
            {
                link.IsEnabled = true;
            }
        }

        // -----------------------------
        // Shorten Path for Display
        // -----------------------------
        private static string ShortenPathForDisplay(string fullPath, int maxLength = 70)
        {
            if (string.IsNullOrWhiteSpace(fullPath) || fullPath.Length <= maxLength)
                return fullPath;

            string fileName = Path.GetFileName(fullPath);
            if (string.IsNullOrWhiteSpace(fileName))
                return fullPath;

            if (fileName.Length >= maxLength - 4)
                return "..." + fileName.Substring(Math.Max(0, fileName.Length - (maxLength - 3)));

            int remaining = maxLength - fileName.Length - 4;
            if (remaining < 10)
                return "...\\" + fileName;

            string start = fullPath.Substring(0, Math.Min(remaining, fullPath.Length));
            return start.TrimEnd('\\') + "\\...\\" + fileName;
        }
    }
}