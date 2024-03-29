﻿using ClosedXML.Excel;
using FaultlessExecution.Abstractions;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSqlServer.Services.ExcelParse
{

    public class ExcelParser : Abstractions.IExcelParser
    {
        private readonly ILogger<ExcelParser> _logger;
        private readonly IFaultlessExecutionService _faultlessExecutionService;

        private class Field
        {
            public int ColumnPosition { get; set; }

            /// <summary>
            /// the raw header/column/field name
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// if stripFieldNameToAlphaAndNumeric is used this will be the stripped down name else it will be the same as <see cref="Name"/>
            /// </summary>
            public string Key { get; set; }
        }

        public ExcelParser(ILogger<ExcelParser> logger, FaultlessExecution.Abstractions.IFaultlessExecutionService faultlessExecutionService)
        {
            _logger = logger;
            _faultlessExecutionService = faultlessExecutionService;
        }

        public ParseResult ParseWorkbook(Stream file, ParseSettings settings)
        {
            var result = new ParseResult();
            using (var workbook = new XLWorkbook(file))
            {
                var sheets = workbook.Worksheets
                    .Where(x => !settings.SkipHiddenWorksheets || x.Visibility == XLWorksheetVisibility.Visible)
                    .Where(x => !settings.ExcludeSheets.Contains(x.Name))
                    .Where(x => !settings.LimitToOnlySheets.Any() || settings.LimitToOnlySheets.Contains(x.Name))
                    .ToList();

                if (!sheets.Any())
                {
                    result.Errors.Add("No worksheets found");
                    return result;
                }
                foreach (var ws in sheets)
                {
                    string workshetName = ws.Name;

                    _logger.LogDebug("start parsing {Worksheet}", workshetName);
                    var worksheetResult = ParseWorksheet(ws, settings);

                    //when pulling in, prefix with worksheetname if they are asking for all worksheets
                    string prefix = settings.AllWorksheets ? $" Worksheet '{workshetName}': " : "";

                    result.Warnings.AddRange(worksheetResult.Warnings.Select(x => $"{prefix}{x}"));
                    result.Errors.AddRange(worksheetResult.Errors.Select(x => $"{prefix}{x}"));
                    result.Records.AddRange(worksheetResult.Records);
                }
            }

            return result;
        }

        private ParseResult ParseWorksheet(IXLWorksheet worksheet, ParseSettings settings)
        {
            var result = new ParseResult();
            var lastRow = worksheet.LastRowUsed();

            // Look for the first row used
            var startingRow = settings.SkipBlankRows ? worksheet.FirstRowUsed() : worksheet.FirstRow();
            if (settings.StartAtRow > 1)
            {
                startingRow = startingRow.RowBelow(settings.StartAtRow - 1);
            }
            int cellPosition = 0;//start at 0 but FIRST thing we'll do is increment; not zero based so that the user gets warnings/messages with a position that makes sense to them
            var fields = new List<Field>();

            _logger.LogDebug("Read Headers...");
            foreach (var col in worksheet.Columns())
            {
                cellPosition++;//yes increment even if Empty
                var cell = startingRow.Cell(cellPosition);


                if (settings.SkipBlankColumns && col.IsEmpty())
                    continue;

                var field = new Field();
                fields.Add(field);
                field.ColumnPosition = cellPosition;

                if (settings.FirstRowIsHeader)
                {
                    field.Name = cell.GetString();
                    field.Key = field.Name;

                    //do field replacement before StripFieldNameToAlphaAndNumeric because they may want to replace something like # with number, which has to be done BEFORE striping thos chars out
                    if (settings.FieldNameReplacements?.Any() == true)
                    {
                        foreach (var replacementSetting in settings.FieldNameReplacements)
                        {
                            field.Key = field.Key.Replace(replacementSetting.MatchingText, replacementSetting.ReplacementText, StringComparison.OrdinalIgnoreCase);
                        }
                    }

                    if (settings.StripFieldNameToAlphaAndNumeric)
                    {
                        field.Key = FieldStrip(field.Key);
                    }

                    if (string.IsNullOrEmpty(field.Key))
                    {
                        field.Key = $"Column{cellPosition}";
                        result.Warnings.Add($"Column Name in Postion {cellPosition} is empty or contains only special characters");
                    }
                    else
                    {
                        if (field.Key.Length > settings.FiedNameCharacterLimit)
                        {
                            if (settings.FieldNameOverLimitSplitFiftyFifty)
                            {
                                //on an even number left and right would be the same, but on an odd number one will be a single char longer, which is why we have to compute right, not just assume both left/right will be the same
                                int leftLength = (int)Math.Round(settings.FiedNameCharacterLimit / 2, decimals: 0, MidpointRounding.AwayFromZero);
                                int rightLength = settings.FiedNameCharacterLimit - leftLength;

                                string leftside = field.Key.Substring(0, leftLength);
                                string rightside = field.Key.Substring(startIndex: field.Key.Length - rightLength, rightLength);

                                string newFieldKey = leftside + rightside;

                                _logger.LogInformation("Field key '{originalFieldKey} was stripped to '{newFieldKey}' in order to meet the {characterLimit} character limit",
                                    field.Key, newFieldKey, settings.FiedNameCharacterLimit);

                                field.Key = newFieldKey;
                            }
                            else
                            {
                                field.Key = field.Key.Substring(0, settings.FiedNameCharacterLimit);
                            }
                        }
                    }
                }
                else
                {
                    field.Name = $"Column{cellPosition}";
                    field.Key = field.Name;
                }

            }

            _logger.LogDebug("Read Rows...");

            //actually build out the import records
            if (fields.Any())
            {
                var row = startingRow;

                if (settings.FirstRowIsHeader)
                {
                    row = row.RowBelow();
                }

                int rowPosition = 0;
                while (row != null)
                {
                    rowPosition++;//yes, increment even it empty
                    _logger.LogDebug("Read Row #{RowNumber}", rowPosition);

                    bool isEmpty = row.IsEmpty();

                    if (isEmpty && !settings.SkipBlankRows)
                    {
                        result.Warnings.Add($"Row {row.RowNumber()} is empty");
                    }

                    if (!isEmpty)
                    {
                        var record = new ImportRecord();
                        result.Records.Add(record);
                        record.RowPosition = rowPosition;
                        record.WorksheetName = worksheet.Name;

                        foreach (var field in fields)
                        {
                            var value = new RecordValue()
                            {
                                FieldKey = field.Key,
                            };

                            //try to get the value; this can be a common point of failure when excel uses functions like "VLOOKUP" and "HLOOKUP" which are not supported by closedxml
                            var cellResult = _faultlessExecutionService.TryExecute(() => value.Value = row.Cell(field.ColumnPosition).GetString());

                            if (!cellResult.WasSuccessful)
                            {
                                var cachedValueResult = _faultlessExecutionService.TryExecute(() => value.Value = row.Cell(field.ColumnPosition).CachedValue?.ToString());
                                if (cachedValueResult.WasSuccessful)
                                {
                                    result.Warnings.Add($"Get value for row {rowPosition} cell {field.ColumnPosition} failed so we used the Cached Value");
                                }
                                else
                                {
                                    result.Errors.Add($"Error getting value for row {rowPosition} cell {field.ColumnPosition} ");
                                    value.Value = "-Error-";
                                }
                            }

                            if (settings.TrimWhiteSpaceFromValues)
                            {
                                value.Value = value.Value?.Trim();
                            }

                            record.Values.Add(value);
                        }
                    }

                    if (row == lastRow)
                    {
                        row = null;
                    }
                    else
                    {
                        row = row.RowBelow();
                    }
                }
            }

            return result;
        }

        private string FieldStrip(string name)
        {
            if (string.IsNullOrEmpty(name))
                return name;

            //keep ONLY a-z and 0-9.  this includes stripping spaces, tabs, crlf, etc
            return System.Text.RegularExpressions.Regex.Replace(name, "[^a-zA-Z0-9]", string.Empty);
        }
    }
}
