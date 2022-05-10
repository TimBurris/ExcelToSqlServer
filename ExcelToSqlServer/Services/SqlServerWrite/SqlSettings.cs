namespace ExcelToSqlServer.Services.SqlServerWrite
{
    public class SqlSettings
    {
        public bool DropTable { get; set; }
        public string SchemaName { get; set; }
        public string TableName { get; set; }
        public string ConnectionString { get; set; }
    }
}
