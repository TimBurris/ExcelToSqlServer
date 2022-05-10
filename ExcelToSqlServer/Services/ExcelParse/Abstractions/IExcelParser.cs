using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSqlServer.Services.ExcelParse.Abstractions
{
    public interface IExcelParser
    {
        ParseResult ParseWorkbook(Stream file, ParseSettings settings);
    }

}
