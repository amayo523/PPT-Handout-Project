using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace PptNotesHandoutMaker.Core
{
    public sealed class BatchPptItem : INotifyPropertyChanged
    {
        private string _pptPath = string.Empty;
        private string _displayFileName = string.Empty;
        private string _pdfTitle = string.Empty;
        private bool _usedFilenameFallback;

        public string PptPath
        {
            get => _pptPath;
            set => SetField(ref _pptPath, value);
        }

        public string DisplayFileName
        {
            get => _displayFileName;
            set => SetField(ref _displayFileName, value);
        }

        public string PdfTitle
        {
            get => _pdfTitle;
            set => SetField(ref _pdfTitle, value);
        }

        public bool UsedFilenameFallback
        {
            get => _usedFilenameFallback;
            set => SetField(ref _usedFilenameFallback, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (Equals(field, value))
                return false;

            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }
}
