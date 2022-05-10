using ExcelToSqlServer.Services.ExcelParse;
using FaultlessExecution.Abstractions;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSqlServer.Services.SqlServerWrite
{

    public class SqlServerWriter : Abstractions.ISqlServerWriter
    {
        private readonly ILogger<SqlServerWriter> _logger;
        private readonly IFaultlessExecutionService _faultlessExecutionService;

        public SqlServerWriter(ILogger<SqlServerWriter> logger, IFaultlessExecutionService faultlessExecutionService)
        {
            _logger = logger;
            _faultlessExecutionService = faultlessExecutionService;
        }

        private void CreateTable(IEnumerable<ImportRecord> records, Microsoft.Data.SqlClient.SqlConnection connection, string qualifiedTableName, string tableName, bool dropTable)
        {
            var columns = records
                              .SelectMany(x => x.Values.Select(y => y.FieldKey))
                              .Distinct(StringComparer.OrdinalIgnoreCase)
                              .ToList();
            string columnSql = string.Join(',', columns.Select(x => $"[{x}] nvarchar(max)"));

            if (dropTable)
            {
                string dropSql = $@"if OBJECT_ID('{qualifiedTableName}', 'U') IS NOT NULL
                                            drop table {qualifiedTableName};";
                Execute(connection, dropSql, parameterValues: null);
            }

            string tableSql = $"CREATE TABLE {qualifiedTableName}([{tableName}Id] uniqueidentifier NOT NULL DEFAULT(newid()), {columnSql})";
            Execute(connection, tableSql, parameterValues: null);

        }

        //TODO: in the future we should tie directly to an ExcelParse ImportRecord, we should have our own and just transfer the data over; that would allow our sqlwriter to be independent of the excel parser
        private void WriteToTable(IEnumerable<ImportRecord> records, Microsoft.Data.SqlClient.SqlConnection connection, string qualifiedTableName)
        {

            /*
                we will build an insert stabment that looks like:
                    Insert into tablename (col1, col2, col3) values(@0, @1, @2);
            
                the values will be passed in a Params so we don't have to worry about any sql injection issues
            */

            foreach (var rec in records)
            {
                //exclude any null/empty values
                var applicableValues = rec.Values.Where(x => !string.IsNullOrEmpty(x.Value)).OrderBy(x => x.FieldKey).ToList();

                //not using the "all columns" because this record might not have many of the columns, so our insert will be only for the columns we have
                string columnNames = string.Join(',', applicableValues.Select(x => $"[{x.FieldKey}]"));

                string values = string.Empty;

                //here is werer we build our "Values" list which will look like "@0,@1,@2,@3" etc
                for (int i = 0; i < applicableValues.Count; i++)
                {
                    if (i > 0)
                    {
                        values = values + ",";
                    }
                    values += $"@{i}";
                }

                string recordSql = $"INSERT INTO {qualifiedTableName}({columnNames}) values({values})";

                //these are the "Real" values, the actual data for @0, @1, etc
                var parameters = applicableValues.Select(x => x.Value).ToList();
                var result = _faultlessExecutionService.TryExecute(() => Execute(connection, recordSql, parameters));

                if (result.WasSuccessful)
                {
                    //TODO: increment a success/fail counter?
                }
            }
        }

        public void WriteToSqlServer(IEnumerable<ImportRecord> records, SqlSettings settings)
        {
            const string worksheetNameReplacementValue = "{WorksheetName}";

            using (var con = new Microsoft.Data.SqlClient.SqlConnection(settings.ConnectionString))
            {
                con.Open();

                //if the want a separate table per worksheet, we'll group all the recrods by their worksheet, else we'll make a single group
                List<IGrouping<string, ImportRecord>>? groupedRecords;
                if (settings.TableName.Contains(worksheetNameReplacementValue, StringComparison.OrdinalIgnoreCase))
                {
                    groupedRecords = records.GroupBy(x => x.WorksheetName).ToList();
                }
                else
                {
                    //this creates one group, with all records in it
                    groupedRecords = records.GroupBy(x => settings.TableName).ToList();
                }

                string schemaName = settings.SchemaName ?? "dbo";

                foreach (var g in groupedRecords)
                {
                    string tableName = settings.TableName.Replace(worksheetNameReplacementValue, g.Key, StringComparison.OrdinalIgnoreCase);
                    string qualifiedTableName = $"[{schemaName}].[{tableName}]";
                    CreateTable(g.ToList(), con, qualifiedTableName, tableName, settings.DropTable);
                    WriteToTable(g.ToList(), con, qualifiedTableName);
                }
            }
        }

        private void Execute(Microsoft.Data.SqlClient.SqlConnection connection, string sql, List<string> parameterValues)
        {
            _logger.LogDebug("Executing {sql} parameters {parameters}", sql, parameterValues);
            var cmd = connection.CreateCommand();
            cmd.CommandText = sql;
            if (parameterValues?.Any() == true)
            {
                int i = 0;
                foreach (var p in parameterValues)
                {
                    var param = cmd.CreateParameter();
                    param.ParameterName = $"@{i}";
                    param.Value = p;
                    param.Size = -1;//-1 means "MAX"
                    param.SqlDbType = System.Data.SqlDbType.NVarChar;
                    i++;

                    cmd.Parameters.Add(param);
                }
            }
            cmd.ExecuteNonQuery();
        }
    }
}
