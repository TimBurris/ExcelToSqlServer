namespace ExcelToSqlServer.Services.ExcelParse
{
    public class ImportRecord
    {
        public List<RecordValue> Values { get; } = new List<RecordValue>();
        public int RowPosition { get; set; }
        public string WorksheetName { get; set; }
    }
}
