namespace ExcelToSqlServer.Services.ExcelParse
{
    public class ParseSettings
    {
        public bool SkipBlankRows { get; set; }
        public bool SkipBlankColumns { get; set; }
        public int StartAtRow { get; set; } = 1;
        public bool FirstRowIsHeader { get; set; } = true;
        public bool SkipHiddenWorksheets { get; set; } = true;
        public bool AllWorksheets { get; set; } = true;
        public bool StripFieldNameToAlphaAndNumeric { get; set; } = true;
        public bool TrimWhiteSpaceFromValues { get; set; } = true;
        public HashSet<string> LimitToOnlySheets { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> ExcludeSheets { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }
}
