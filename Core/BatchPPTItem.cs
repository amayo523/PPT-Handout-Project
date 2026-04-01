namespace PptNotesHandoutMaker.Core
{
    public sealed class BatchPptItem
    {
        public string PptPath { get; set; } = "";
        public string PdfTitle { get; set; } = "";
        public bool UsedFilenameFallback { get; set; }
    }
}