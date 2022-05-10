namespace ExcelToSqlServer.Services.ExcelParse
{
    public class ParseResult
    {
        public List<ImportRecord> Records { get; } = new List<ImportRecord>();
        public List<string> Errors { get; } = new List<string>();
        public List<string> Warnings { get; } = new List<string>();
    }
}
