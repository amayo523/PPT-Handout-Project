using PptNotesHandoutMaker.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PptNotesHandoutMaker
{
    public partial class MainWindow : Window
    {
        // -----------------------------
        // Debug / feature flags
        // -----------------------------
        private static readonly bool FORCE_FILENAME_FALLBACK_FOR_TEST = false;
        private const bool ENABLE_SHARED_DRIVE_WARNING = false;
        private static readonly bool ENABLE_TIMING_LOGS = true;

        // -----------------------------
        // State
        // -----------------------------
        private bool _isGenerating;
        private bool _isReadingTitles;
        private bool _isLoaded;

        private readonly ObservableCollection<BatchPptItem> _selectedPpts = new();
        private readonly HashSet<string> _selectedPptPaths = new(StringComparer.OrdinalIgnoreCase);
        private static readonly Regex DigitsOnlyRegex = new(@"^\d+$");
        private static readonly Regex MonthYearRegex = new(@"^(0[1-9]|1[0-2])\/\d{4}$");
        private const string DatePlaceholder = "MM/YYYY";

        // -----------------------------
        // Construction
        // -----------------------------
        public MainWindow()
        {
            InitializeComponent();

            BatchItemsControl.ItemsSource = _selectedPpts;

            if (UseCurrentDateCheckBox.IsChecked == true)
            {
                DateBox.Text = DateTime.Now.ToString("MM/yyyy");
                DateBox.Foreground = System.Windows.Media.Brushes.Black;
            }
            else
            {
                DateBox.Text = DatePlaceholder;
                DateBox.Foreground = System.Windows.Media.Brushes.Gray;
            }

            UpdateDateInputState();

            UpdatePptCountDisplay();
            UpdateDropHintVisibility();
            UpdateDropZoneVisual(DropZoneState.Normal);
            UpdateUiState();

            _isLoaded = true;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _isLoaded = true;

            UpdatePptCountDisplay();
            UpdateDropHintVisibility();
            UpdateDropZoneVisual(DropZoneState.Normal);
            UpdateDateInputState();
            UpdateUiState();
        }

        // -----------------------------
        // Event handlers
        // -----------------------------
        private void AnyOptionChanged(object sender, RoutedEventArgs e)
        {
            if (!_isLoaded)
                return;

            UpdateUiState();
        }

        private void AnyInputChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isLoaded)
                return;

            UpdateUiState();
        }

        private async void RetryTitleRead_Link_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not System.Windows.Documents.Hyperlink link ||
                link.Tag is not BatchPptItem item)
            {
                return;
            }

            string pptPath = (item.PptPath ?? "").Trim();

            if (string.IsNullOrWhiteSpace(pptPath) || !File.Exists(pptPath))
            {
                AppendStatus("Retry failed: file not found.");
                AppendStatus(pptPath);
                return;
            }

            string displayName = GetShortDisplayName(pptPath);
            Stopwatch? retryTimer = StartTiming();

            try
            {
                link.IsEnabled = false;

                StatusLabel.Text = $"Status: Retrying title read for {displayName}";
                AppendStatus("--------------------------------------------------");
                AppendStatus($"Retrying title read for: {pptPath}");

                string? retriedTitle = await StaTask.Run(() =>
                    PowerPointTitleReader.TryReadFirstSlideTitle(pptPath));

                StopTiming(retryTimer);

                if (!string.IsNullOrWhiteSpace(retriedTitle))
                {
                    item.PdfTitle = retriedTitle.Trim();
                    item.UsedFilenameFallback = false;

                    AppendStatus($"Retry succeeded. Updated PDF title to: {item.PdfTitle}");
                    AppendTiming($"Retry title read time: {FormatElapsed(retryTimer?.Elapsed)}");
                    StatusLabel.Text = $"Status: Retry succeeded for {displayName}";
                }
                else
                {
                    AppendStatus("Retry did not find a usable slide title.");
                    AppendTiming($"Retry title read time: {FormatElapsed(retryTimer?.Elapsed)}");
                    StatusLabel.Text = $"Status: Retry found no title for {displayName}";
                }

                RefreshSelectionUi();
            }
            catch (Exception ex)
            {
                StopTiming(retryTimer);
                AppendStatus($"Retry failed for: {pptPath}");
                AppendStatus(ex.Message);
                AppendTiming($"Retry time before failure: {FormatElapsed(retryTimer?.Elapsed)}");
                StatusLabel.Text = $"Status: Retry failed for {displayName}";
            }
            finally
            {
                link.IsEnabled = true;
            }
        }

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

            UpdateDropZoneVisual(allArePowerPoints
                ? DropZoneState.Valid
                : DropZoneState.Invalid);

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
                    string firstDir = Path.GetDirectoryName(_selectedPpts[0].PptPath) ?? string.Empty;
                    if (Directory.Exists(firstDir))
                        dlg.InitialDirectory = firstDir;
                }
                catch
                {
                }
            }

            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK ||
                string.IsNullOrWhiteSpace(dlg.SelectedPath))
            {
                return;
            }

            OutPathBox.Text = dlg.SelectedPath;

            AppendStatus("Output folder:");
            AppendStatus(dlg.SelectedPath);

            UpdateUiState();
        }

        private async void Generate_Click(object sender, RoutedEventArgs e)
        {
            if (_isGenerating || _isReadingTitles)
                return;

            string rawClassName = (ClassNameBox.Text ?? "").Trim();
            string rawRevision = (RevisionBox.Text ?? "").Trim();
            string className = BuildDisplayClassName();
            string displayDate = BuildDisplayDate();
            string outputFolder = (OutPathBox.Text ?? "").Trim();
            bool alwaysUseTempLocalCopy = AlwaysUseTempCopyCheckBox.IsChecked == true;

            if (_selectedPpts.Count == 0)
            {
                AppendStatus("No PowerPoint files selected.");
                UpdateUiState();
                return;
            }

            // Hard-stop validation
            if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Please choose a valid output folder before generating.",
                    "Missing Output Folder",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }

            // Warning-only validation
            var missingItems = new List<string>();

            if (string.IsNullOrWhiteSpace(rawClassName))
                missingItems.Add("Class name");

            if (string.IsNullOrWhiteSpace(rawRevision))
                missingItems.Add("Revision");

            if (UseCurrentDateCheckBox.IsChecked != true && !MonthYearRegex.IsMatch(displayDate))
                missingItems.Add("Date (MM/YYYY)");

            bool hasBlankTitles = _selectedPpts.Any(x => string.IsNullOrWhiteSpace(x.PdfTitle));
            if (hasBlankTitles)
                missingItems.Add("One or more PDF titles");

            if (missingItems.Count > 0)
            {
                string warningMessage =
                    "The following information is missing:\n\n• " +
                    string.Join("\n• ", missingItems) +
                    "\n\nDo you want to generate the PDF(s) anyway?";

                var result = System.Windows.MessageBox.Show(
                    this,
                    warningMessage,
                    "Generate With Missing Information?",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning,
                    MessageBoxResult.No);

                if (result != MessageBoxResult.Yes)
                    return;
            }

            foreach (var item in _selectedPpts)
            {
                if (string.IsNullOrWhiteSpace(item.PptPath) || !File.Exists(item.PptPath))
                {
                    AppendStatus($"File not found: {item.PptPath}");
                    UpdateUiState();
                    return;
                }
            }

            StatusBox.Clear();

            Stopwatch? batchGenerateTimer = StartTiming();

            _isGenerating = true;
            StatusLabel.Text = $"Status: Starting batch (0 of {_selectedPpts.Count})";
            UpdateUiState();

            IProgress<string> progress = new Progress<string>(msg =>
            {
                if (msg.StartsWith("SLIDE_PROGRESS|", StringComparison.Ordinal))
                {
                    var parts = msg.Split('|');
                    if (parts.Length >= 3 &&
                        int.TryParse(parts[1], out int current) &&
                        int.TryParse(parts[2], out int total))
                    {
                        StatusLabel.Text = $"Status: Exporting slide {current}/{total}...";
                    }

                    return;
                }

                AppendStatus(msg);
            });

            try
            {
                progress.Report($"Starting batch generation for {_selectedPpts.Count} file(s)...");

                await StaTask.Run(() =>
                {
                    for (int i = 0; i < _selectedPpts.Count; i++)
                    {
                        var item = _selectedPpts[i];
                        string pptPath = (item.PptPath ?? string.Empty).Trim();
                        string pdfTitle = (item.PdfTitle ?? string.Empty).Trim();

                        int currentIndex = i + 1;
                        int totalCount = _selectedPpts.Count;
                        string displayName = GetShortDisplayName(pptPath);

                        Dispatcher.Invoke(() =>
                        {
                            StatusLabel.Text = $"Status: Processing ({currentIndex} of {totalCount}) - {displayName}";
                        });

                        Stopwatch? fileGenerateTimer = StartTiming();

                        try
                        {
                            progress.Report("--------------------------------------------------");
                            progress.Report($"Processing: {Path.GetFileName(pptPath)}");

                            if (string.IsNullOrWhiteSpace(pptPath) || !File.Exists(pptPath))
                            {
                                StopTiming(fileGenerateTimer);
                                progress.Report("Skipped: file not found.");
                                ReportTiming(progress, $"Time for skipped file: {FormatElapsed(fileGenerateTimer?.Elapsed)}");
                                continue;
                            }

                            if (string.IsNullOrWhiteSpace(pdfTitle))
                                pdfTitle = Path.GetFileNameWithoutExtension(pptPath);

                            string outputPdfPath = Path.Combine(
                                outputFolder,
                                $"{Path.GetFileNameWithoutExtension(pptPath)}_Instructor Guide.pdf");

                            outputPdfPath = GetNextAvailableFilePath(outputPdfPath);

                            progress.Report($"PDF title: {pdfTitle}");
                            progress.Report($"Output: {outputPdfPath}");

                            var opt = new HandoutOptions
                            {
                                ClassName = className,
                                PdfTitle = pdfTitle,
                                CreatedDate = displayDate,
                                ShowNoNotesPlaceholder = true,
                                AlwaysUseTempLocalCopy = alwaysUseTempLocalCopy
                            };

                            var gen = new HandoutGenerator(opt);
                            gen.Generate(pptPath, outputPdfPath, progress);

                            StopTiming(fileGenerateTimer);
                            progress.Report($"Finished: {outputPdfPath}");
                            ReportTiming(progress, $"Time for {Path.GetFileName(pptPath)}: {FormatElapsed(fileGenerateTimer?.Elapsed)}");
                        }
                        catch (Exception ex)
                        {
                            StopTiming(fileGenerateTimer);
                            progress.Report("ERROR processing file:");
                            progress.Report(pptPath);
                            progress.Report(ex.ToString());
                            ReportTiming(progress, $"Time before failure for {Path.GetFileName(pptPath)}: {FormatElapsed(fileGenerateTimer?.Elapsed)}");
                        }
                    }
                });

                StopTiming(batchGenerateTimer);
                AppendTiming($"PDF batch generation completed in {FormatElapsed(batchGenerateTimer?.Elapsed)}");
                StatusLabel.Text = "Status: Batch complete";

                Process.Start(new ProcessStartInfo(outputFolder)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                StopTiming(batchGenerateTimer);
                progress.Report("BATCH ERROR:");
                progress.Report(ex.ToString());
                AppendTiming($"PDF batch generation stopped after {FormatElapsed(batchGenerateTimer?.Elapsed)}");
                StatusLabel.Text = "Status: Batch error";
            }
            finally
            {
                _isGenerating = false;
                UpdateUiState();
            }
        }

        private void RemoveBatchItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not System.Windows.Controls.Button btn || btn.Tag is not BatchPptItem item)
                return;

            _selectedPpts.Remove(item);
            _selectedPptPaths.Remove(item.PptPath);

            AppendStatus($"Removed: {Path.GetFileName(item.PptPath)}");
            RefreshSelectionUi();
        }

        private void ClearBatch_Click(object sender, RoutedEventArgs e)
        {
            _selectedPpts.Clear();
            _selectedPptPaths.Clear();

            StatusBox.Clear();
            AppendStatus("Batch list cleared.");

            RefreshSelectionUi();
        }

        // -----------------------------
        // Core workflows
        // -----------------------------
        private async Task AddPowerPointsAsync(IEnumerable<string> filePaths)
        {
            if (filePaths == null)
                return;

            var paths = filePaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .ToList();

            if (paths.Count == 0)
                return;

            Stopwatch? batchReadTimer = StartTiming();

            string lastSelectedPath = paths[^1];
            string? lastSelectedFolder = Path.GetDirectoryName(lastSelectedPath);

            if (!string.IsNullOrWhiteSpace(lastSelectedFolder) && Directory.Exists(lastSelectedFolder))
                OutPathBox.Text = lastSelectedFolder;

            _isReadingTitles = true;
            UpdateUiState();
            UpdateDropHintVisibility();

            try
            {
                AppendStatus("--------------------------------------------------");
                AppendStatus("Starting title-read batch...");
                StatusLabel.Text = $"Status: Starting title read (0 of {paths.Count})";

                for (int i = 0; i < paths.Count; i++)
                {
                    string selectedPath = paths[i];

                    if (ENABLE_SHARED_DRIVE_WARNING && !ConfirmSharedDriveWarning(this, selectedPath))
                        continue;

                    if (!_selectedPptPaths.Add(selectedPath))
                    {
                        AppendStatus($"Skipped already-added file: {selectedPath}");
                        continue;
                    }

                    string displayName = GetShortDisplayName(selectedPath);
                    StatusLabel.Text = $"Status: Reading titles ({i + 1} of {paths.Count}) - {displayName}";

                    Stopwatch? fileReadTimer = StartTiming();

                    string detectedTitle;
                    bool usedFallback = false;

                    if (FORCE_FILENAME_FALLBACK_FOR_TEST)
                    {
                        detectedTitle = Path.GetFileNameWithoutExtension(selectedPath);
                        usedFallback = true;
                        StopTiming(fileReadTimer);
                        AppendStatus($"TEST MODE: forcing filename fallback for {selectedPath}");
                    }
                    else
                    {
                        try
                        {
                            detectedTitle = await StaTask.Run(() =>
                                PowerPointTitleReader.TryReadFirstSlideTitle(selectedPath)) ?? string.Empty;
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

                        StopTiming(fileReadTimer);
                    }

                    _selectedPpts.Add(new BatchPptItem
                    {
                        PptPath = selectedPath,
                        DisplayFileName = displayName,
                        PdfTitle = detectedTitle,
                        UsedFilenameFallback = usedFallback
                    });

                    RefreshSelectionUi();

                    AppendStatus($"Added: {selectedPath}");
                    AppendTiming($"Title read time for {Path.GetFileName(selectedPath)}: {FormatElapsed(fileReadTimer?.Elapsed)}");
                }

                StopTiming(batchReadTimer);
                AppendTiming($"Title read + batch population completed in {FormatElapsed(batchReadTimer?.Elapsed)}");
                StatusLabel.Text = $"Status: Added {_selectedPpts.Count} PowerPoint file(s)";
                RefreshSelectionUi();
            }
            finally
            {
                if (batchReadTimer?.IsRunning == true)
                {
                    StopTiming(batchReadTimer);
                    AppendTiming($"Title read + batch population completed in {FormatElapsed(batchReadTimer?.Elapsed)}");
                }

                _isReadingTitles = false;
                UpdateUiState();
            }
        }

        // -----------------------------
        // UI refresh helpers
        // -----------------------------
        private void RefreshSelectionUi()
        {
            UpdatePptCountDisplay();
            UpdateDropHintVisibility();
            UpdateUiState();
        }

        private void UpdatePptCountDisplay()
        {
            int count = _selectedPpts.Count;

            if (count > 0)
            {
                PptCountText.Text = count == 1 ? "1 file selected" : $"{count} files selected";
                PptCountText.Visibility = Visibility.Visible;
            }
            else
            {
                PptCountText.Text = string.Empty;
                PptCountText.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdateDropHintVisibility()
        {
            DropHintText.Visibility = _selectedPpts.Count == 0
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        private void UpdateUiState()
        {
            bool isBusy = _isGenerating || _isReadingTitles;
            bool hasSelectedPpts = _selectedPpts.Count > 0;

            GenerateBtn.IsEnabled = !isBusy && hasSelectedPpts;
            AddPptBtn.IsEnabled = !isBusy;
            ClearBatchBtn.IsEnabled = !isBusy && hasSelectedPpts;
        }

        private void UpdateDropZoneVisual(DropZoneState state)
        {
            switch (state)
            {
                case DropZoneState.Normal:
                    DropZoneBorder.BorderBrush = System.Windows.Media.Brushes.LightGray;
                    DropZoneBorder.Background = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(250, 250, 250));
                    break;

                case DropZoneState.Valid:
                    DropZoneBorder.BorderBrush = System.Windows.Media.Brushes.ForestGreen;
                    DropZoneBorder.Background = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(240, 255, 240));
                    break;

                case DropZoneState.Invalid:
                    DropZoneBorder.BorderBrush = System.Windows.Media.Brushes.IndianRed;
                    DropZoneBorder.Background = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(255, 240, 240));
                    break;
            }
        }

        // -----------------------------
        // Class name + revision helper
        // -----------------------------
        private string BuildDisplayClassName()
        {
            string className = (ClassNameBox.Text ?? "").Trim();
            string revision = (RevisionBox.Text ?? "").Trim();

            if (string.IsNullOrWhiteSpace(className))
                return "";

            if (string.IsNullOrWhiteSpace(revision))
                return className;

            return $"{className}, Rev. {revision}";
        }
        private void RevisionBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !DigitsOnlyRegex.IsMatch(e.Text);
        }
        private void RevisionBox_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (!e.DataObject.GetDataPresent(typeof(string)))
            {
                e.CancelCommand();
                return;
            }

            string pastedText = (string)e.DataObject.GetData(typeof(string))!;

            if (string.IsNullOrWhiteSpace(pastedText) || !DigitsOnlyRegex.IsMatch(pastedText))
                e.CancelCommand();
        }

        // -----------------------------
        // Date input helper
        // -----------------------------
        private void UseCurrentDateCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isLoaded)
                return;

            if (UseCurrentDateCheckBox.IsChecked == true)
            {
                DateBox.Text = DateTime.Now.ToString("MM/yyyy");
                DateBox.Foreground = System.Windows.Media.Brushes.Black;
            }
            else
            {
                DateBox.Text = DatePlaceholder;
                DateBox.Foreground = System.Windows.Media.Brushes.Gray;
                DateBox.CaretIndex = 0;
            }

            UpdateDateInputState();
            UpdateUiState();
        }

        private void UpdateDateInputState()
        {
            if (DateBox == null || UseCurrentDateCheckBox == null)
                return;

            bool useCurrentDate = UseCurrentDateCheckBox.IsChecked == true;
            DateBox.IsEnabled = !useCurrentDate;
        }

        private string BuildDisplayDate()
        {
            if (UseCurrentDateCheckBox.IsChecked == true)
                return DateTime.Now.ToString("MM/yyyy");

            string manualDate = (DateBox.Text ?? "").Trim();

            if (string.Equals(manualDate, DatePlaceholder, StringComparison.OrdinalIgnoreCase))
                return "";

            return manualDate;
        }

        // -----------------------------
        // Date textbox input restrictions helper
        // -----------------------------
        private void DateBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (sender is not System.Windows.Controls.TextBox textBox)
                return;

            if (UseCurrentDateCheckBox.IsChecked == true)
            {
                e.Handled = true;
                return;
            }

            if (string.IsNullOrWhiteSpace(e.Text) || !e.Text.All(char.IsDigit))
            {
                e.Handled = true;
                return;
            }

            string digits = ExtractDateDigits(textBox.Text);
            int digitCaret = GetDigitCaretIndex(textBox.Text, textBox.CaretIndex);

            foreach (char ch in e.Text)
            {
                if (digits.Length >= 6)
                    break;

                digits = digits.Insert(Math.Min(digitCaret, digits.Length), ch.ToString());
                digitCaret++;
            }

            digits = digits.Length > 6 ? digits.Substring(0, 6) : digits;

            textBox.Text = BuildMaskedDateText(digits);
            textBox.Foreground = digits.Length == 0
                ? System.Windows.Media.Brushes.Gray
                : System.Windows.Media.Brushes.Black;

            SetDateCaretFromDigitIndex(textBox, digitCaret);

            e.Handled = true;
        }

        private void DateBox_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (sender is not System.Windows.Controls.TextBox textBox)
                return;

            if (UseCurrentDateCheckBox.IsChecked == true)
                return;

            if (e.Key == Key.Space)
            {
                e.Handled = true;
                return;
            }

            if (e.Key != Key.Back && e.Key != Key.Delete)
                return;

            string digits = ExtractDateDigits(textBox.Text);
            int digitCaret = GetDigitCaretIndex(textBox.Text, textBox.CaretIndex);

            if (textBox.SelectionLength > 0)
            {
                int selStartDigit = GetDigitCaretIndex(textBox.Text, textBox.SelectionStart);
                int selEndDigit = GetDigitCaretIndex(textBox.Text, textBox.SelectionStart + textBox.SelectionLength);

                if (selEndDigit > selStartDigit && selStartDigit < digits.Length)
                {
                    digits = digits.Remove(selStartDigit, Math.Min(selEndDigit - selStartDigit, digits.Length - selStartDigit));
                    textBox.Text = BuildMaskedDateText(digits);
                    textBox.Foreground = digits.Length == 0
                        ? System.Windows.Media.Brushes.Gray
                        : System.Windows.Media.Brushes.Black;
                    SetDateCaretFromDigitIndex(textBox, selStartDigit);
                }

                e.Handled = true;
                return;
            }

            if (e.Key == Key.Back)
            {
                if (digitCaret > 0 && digits.Length > 0)
                {
                    digits = digits.Remove(digitCaret - 1, 1);
                    digitCaret--;
                }
            }
            else if (e.Key == Key.Delete)
            {
                if (digitCaret < digits.Length && digits.Length > 0)
                {
                    digits = digits.Remove(digitCaret, 1);
                }
            }

            textBox.Text = BuildMaskedDateText(digits);
            textBox.Foreground = digits.Length == 0
                ? System.Windows.Media.Brushes.Gray
                : System.Windows.Media.Brushes.Black;

            SetDateCaretFromDigitIndex(textBox, digitCaret);

            e.Handled = true;
        }

        private void DateBox_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (sender is not System.Windows.Controls.TextBox textBox)
            {
                e.CancelCommand();
                return;
            }

            if (UseCurrentDateCheckBox.IsChecked == true)
            {
                e.CancelCommand();
                return;
            }

            if (!e.DataObject.GetDataPresent(typeof(string)))
            {
                e.CancelCommand();
                return;
            }

            string pastedText = ((string)e.DataObject.GetData(typeof(string))!) ?? "";
            string pastedDigits = new string(pastedText.Where(char.IsDigit).ToArray());

            if (pastedDigits.Length == 0)
            {
                e.CancelCommand();
                return;
            }

            string digits = ExtractDateDigits(textBox.Text);
            int selStartDigit = GetDigitCaretIndex(textBox.Text, textBox.SelectionStart);
            int selEndDigit = GetDigitCaretIndex(textBox.Text, textBox.SelectionStart + textBox.SelectionLength);

            if (selEndDigit > selStartDigit && selStartDigit < digits.Length)
            {
                digits = digits.Remove(selStartDigit, Math.Min(selEndDigit - selStartDigit, digits.Length - selStartDigit));
            }

            int insertIndex = selStartDigit;
            foreach (char ch in pastedDigits)
            {
                if (digits.Length >= 6)
                    break;

                digits = digits.Insert(Math.Min(insertIndex, digits.Length), ch.ToString());
                insertIndex++;
            }

            textBox.Text = BuildMaskedDateText(digits);
            textBox.Foreground = digits.Length == 0
                ? System.Windows.Media.Brushes.Gray
                : System.Windows.Media.Brushes.Black;

            SetDateCaretFromDigitIndex(textBox, insertIndex);

            e.CancelCommand();
        }

        private static string ExtractDateDigits(string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            return new string(text.Where(char.IsDigit).ToArray());
        }

        private static string BuildMaskedDateText(string digits)
        {
            char[] chars = DatePlaceholder.ToCharArray();

            for (int i = 0; i < Math.Min(digits.Length, 6); i++)
            {
                int targetIndex = i < 2 ? i : i + 1; // skip slash
                chars[targetIndex] = digits[i];
            }

            return new string(chars);
        }

        private static int GetDigitCaretIndex(string? text, int caretIndex)
        {
            string safeText = text ?? string.Empty;

            int count = 0;
            int limit = Math.Min(caretIndex, safeText.Length);

            for (int i = 0; i < limit; i++)
            {
                if (char.IsDigit(safeText[i]))
                    count++;
            }

            return count;
        }

        private static void SetDateCaretFromDigitIndex(System.Windows.Controls.TextBox textBox, int digitIndex)
        {
            digitIndex = Math.Max(0, Math.Min(6, digitIndex));

            int caretIndex = digitIndex switch
            {
                0 => 0,
                1 => 1,
                2 => 3,
                3 => 4,
                4 => 5,
                5 => 6,
                _ => 7
            };

            textBox.CaretIndex = caretIndex;
        }

        private void DateBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (UseCurrentDateCheckBox.IsChecked == true)
                return;

            string digits = ExtractDateDigits(DateBox.Text);
            SetDateCaretFromDigitIndex(DateBox, digits.Length);
        }

        // -----------------------------
        // Status helpers
        // -----------------------------
        private void AppendStatus(string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => AppendStatus(message));
                return;
            }

            StatusBox.AppendText(message + Environment.NewLine);
            StatusBox.ScrollToEnd();
        }

        private void AppendTiming(string message)
        {
            if (!ENABLE_TIMING_LOGS)
                return;

            AppendStatus(message);
        }

        private static Stopwatch? StartTiming()
        {
            return ENABLE_TIMING_LOGS ? Stopwatch.StartNew() : null;
        }

        private static void StopTiming(Stopwatch? stopwatch)
        {
            if (stopwatch?.IsRunning == true)
                stopwatch.Stop();
        }

        private static void ReportTiming(IProgress<string> progress, string message)
        {
            if (!ENABLE_TIMING_LOGS)
                return;

            progress.Report(message);
        }

        private static string FormatElapsed(TimeSpan? elapsed)
        {
            if (elapsed == null)
                return string.Empty;

            if (elapsed.Value.TotalHours >= 1)
                return elapsed.Value.ToString(@"h\:mm\:ss\.ff");

            if (elapsed.Value.TotalMinutes >= 1)
                return elapsed.Value.ToString(@"m\:ss\.ff");

            return elapsed.Value.ToString(@"s\.ff") + " sec";
        }

        // -----------------------------
        // File/path helpers
        // -----------------------------
        private static bool IsPowerPointFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            string ext = Path.GetExtension(path);
            return ext.Equals(".pptx", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".ppt", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetShortDisplayName(string path, int maxLength = 60)
        {
            string fileName = Path.GetFileName(path);

            if (string.IsNullOrWhiteSpace(fileName) || fileName.Length <= maxLength)
                return fileName;

            return fileName[..(maxLength - 3)] + "...";
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
            driveLabel = string.Empty;

            if (string.IsNullOrWhiteSpace(filePath))
                return false;

            if (filePath.StartsWith(@"\\", StringComparison.Ordinal))
            {
                driveLabel = TryGetUncShareRoot(filePath) ?? "network share";
                return true;
            }

            try
            {
                string root = Path.GetPathRoot(filePath) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(root))
                    return false;

                var di = new DriveInfo(root);
                if (di.DriveType == DriveType.Network)
                {
                    driveLabel = root.TrimEnd('\\');
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static string? TryGetUncShareRoot(string uncPath)
        {
            try
            {
                var parts = uncPath.Split('\\', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2)
                    return $@"\\{parts[0]}\{parts[1]}";
            }
            catch
            {
            }

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

        private enum DropZoneState
        {
            Normal,
            Valid,
            Invalid
        }
    }
}