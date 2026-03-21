namespace SheetAppendApp
{
    public sealed class MergeOptions
    {
        public bool AutoRepairBrokenReferences { get; set; } = true;

        // Khuyến nghị: thao tác trên bản copy temp để không “bẩn” file gốc
        public bool WorkOnTempCopy { get; set; } = true;

        // true = copy cả style + value, false = values only (nhanh hơn)
        public bool CopyStyles { get; set; } = true;

        // Sheet index để lấy header/base, mặc định Sheet 1
        public int DestinationSheetIndex { get; set; } = 1;

        // Bỏ header ở các sheet nguồn (dòng đầu UsedRange)
        public bool SkipHeaderOnSourceSheets { get; set; } = true;
    }
}