using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace PptNotesHandoutMaker.Core
{
    public sealed class HandoutOptions
    {
        public PageSize PageSize { get; set; } = PageSizes.Letter;

        public float PageMarginPts { get; set; } = 24f;
        public float GapPts { get; set; } = 12f;

        public float LeftColumnRel { get; set; } = 1.0f;
        public float RightColumnRel { get; set; } = 1.2f;

        public float NotesFontSize { get; set; } = 9f;
        public int SlideExportWidthPx { get; set; } = 1200;

        public bool ShowNoNotesPlaceholder { get; set; } = true;
        public bool SkipSlidesWithNoNotes { get; set; } = false;
        public bool AlwaysUseTempLocalCopy { get; set; } = false;

        public string? ClassName { get; set; }
        public string? PdfTitle { get; set; }
        public string? CreatedDate { get; set; }

    }
}
