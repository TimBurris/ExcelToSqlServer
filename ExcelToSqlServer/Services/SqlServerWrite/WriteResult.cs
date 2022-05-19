namespace ExcelToSqlServer.Services.SqlServerWrite
{
    public class WriteResult
    {
        public List<string> Errors { get; } = new List<string>();
        public List<string> Warnings { get; } = new List<string>();
    }
}
