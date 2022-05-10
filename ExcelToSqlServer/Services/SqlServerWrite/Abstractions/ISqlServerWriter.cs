using ExcelToSqlServer.Services.ExcelParse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSqlServer.Services.SqlServerWrite.Abstractions
{
    public interface ISqlServerWriter
    {
        void WriteToSqlServer(IEnumerable<ImportRecord> records, SqlSettings settings);
    }
}
