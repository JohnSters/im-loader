using System.Collections.Generic;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System;
using System.Linq;

namespace IMLoader
{
    public static class ExcelHelper
    {
        private static bool TryParseDate(IXLCell cell, out DateTime date)
        {
            date = default;
            if (cell.DataType == XLDataType.DateTime)
            {
                date = cell.GetDateTime();
                return true;
            }
            string dateString = cell.GetString().Trim();
            if (string.IsNullOrWhiteSpace(dateString)) return false;

            string[] formats = { "yyyy/MM/dd", "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd-MM-yyyy", "MM-dd-yyyy", "M/d/yyyy" };
            if (DateTime.TryParseExact(dateString, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out date))
            {
                return true;
            }
            return DateTime.TryParse(dateString, out date);
        }

        public static List<string> GetSheetNames(string filePath)
        {
            var sheetNames = new List<string>();
            using (var workbook = new XLWorkbook(filePath))
            {
                foreach (var ws in workbook.Worksheets)
                {
                    sheetNames.Add(ws.Name);
                }
            }
            return sheetNames;
        }

        public static List<string> GetHeaders(string filePath, string sheetName)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                if (!workbook.Worksheets.TryGetWorksheet(sheetName, out var ws) || ws == null)
                    throw new ArgumentException($"Sheet '{sheetName}' not found in file '{filePath}'.");
                var headers = new List<string>();
                foreach (var cell in ws.Row(1).CellsUsed())
                {
                    headers.Add(cell.GetString());
                }
                return headers;
            }
        }

        public static string ExtractUnitNumberFromFileName(string filePath)
        {
            var fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            var match = Regex.Match(fileName, @"CT0*([1-9][0-9]*)", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;
            return string.Empty;
        }

        public static void MergeFiles(
            string masterFilePath,
            string masterSheet,
            List<(string filePath, string sheetName)> filesToMerge,
            string outputFilePath)
        {
            using var masterWb = new XLWorkbook(masterFilePath);
            if (!masterWb.Worksheets.TryGetWorksheet(masterSheet, out var masterWs) || masterWs == null)
                throw new ArgumentException($"Sheet '{masterSheet}' not found in master file.");
            var masterHeaders = masterWs.Row(1).CellsUsed().Select(c => c.GetString()).ToList();

            int masterNextDateIdx = masterHeaders.FindIndex(h => h.Trim().Equals("Next Date", StringComparison.OrdinalIgnoreCase));
            int masterReoccurringIdx = masterHeaders.FindIndex(h => h.Trim().Equals("Reoccurring", StringComparison.OrdinalIgnoreCase));
            int masterLastDateIdx = masterHeaders.FindIndex(h => h.Trim().Equals("Last Date", StringComparison.OrdinalIgnoreCase));
            int masterIntervalIdx = masterHeaders.FindIndex(h => h.Trim().Equals("Desired Interval", StringComparison.OrdinalIgnoreCase));

            int masterColCount = masterHeaders.Count;
            var lastRow = masterWs.LastRowUsed();
            int masterLastRow = lastRow != null ? lastRow.RowNumber() : 1; // 1 = header row, so data starts at 2

            foreach (var (filePath, sheetName) in filesToMerge)
            {
                using var wb = new XLWorkbook(filePath);
                if (!wb.Worksheets.TryGetWorksheet(sheetName, out var ws) || ws == null)
                    throw new ArgumentException($"Sheet '{sheetName}' not found in file '{filePath}'.");
                var headers = ws.Row(1).CellsUsed().Select(c => c.GetString()).ToList();

                // New: Build header map with partial matching
                var headerMap = new List<int>();
                for (int i = 0; i < masterHeaders.Count; i++)
                {
                    string masterHeader = masterHeaders[i].Trim();
                    int exactIdx = headers.FindIndex(h => string.Equals(h.Trim(), masterHeader, StringComparison.OrdinalIgnoreCase));
                    if (exactIdx >= 0)
                    {
                        headerMap.Add(exactIdx);
                        continue;
                    }
                    // Partial match: find the merge header that is a substring of the master header or vice versa, prefer the longest match
                    int bestIdx = -1;
                    int bestLength = 0;
                    for (int j = 0; j < headers.Count; j++)
                    {
                        string mergeHeader = headers[j].Trim();
                        if (masterHeader.IndexOf(mergeHeader, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            mergeHeader.IndexOf(masterHeader, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            int matchLength = Math.Max(masterHeader.Length, mergeHeader.Length);
                            if (matchLength > bestLength)
                            {
                                bestLength = matchLength;
                                bestIdx = j;
                            }
                        }
                    }
                    headerMap.Add(bestIdx);
                }
                string unitNumber = ExtractUnitNumberFromFileName(filePath);
                int row = 3;
                while (true)
                {
                    var dataRow = ws.Row(row);
                    if (dataRow.IsEmpty()) break;

                    DateTime? calculatedNextDate = null;

                    if (masterNextDateIdx != -1 && masterReoccurringIdx != -1 && masterLastDateIdx != -1 && masterIntervalIdx != -1)
                    {
                        int srcNextDateIdx = headerMap[masterNextDateIdx];
                        int srcReoccurringIdx = headerMap[masterReoccurringIdx];
                        int srcLastDateIdx = headerMap[masterLastDateIdx];
                        int srcIntervalIdx = headerMap[masterIntervalIdx];

                        if (srcReoccurringIdx != -1 && srcLastDateIdx != -1 && srcIntervalIdx != -1)
                        {
                            var nextDateCell = (srcNextDateIdx != -1) ? dataRow.Cell(srcNextDateIdx + 1) : null;
                            var reoccurringCell = dataRow.Cell(srcReoccurringIdx + 1);

                            bool isReoccurring = false;
                            if (reoccurringCell.DataType == XLDataType.Boolean) isReoccurring = reoccurringCell.GetBoolean();
                            else if (reoccurringCell.DataType == XLDataType.Text) bool.TryParse(reoccurringCell.GetString(), out isReoccurring);

                            if (!isReoccurring && (nextDateCell == null || nextDateCell.IsEmpty()))
                            {
                                var lastDateCell = dataRow.Cell(srcLastDateIdx + 1);
                                var intervalCell = dataRow.Cell(srcIntervalIdx + 1);
                                if (TryParseDate(lastDateCell, out DateTime lastDate) && intervalCell.TryGetValue(out double intervalMonths))
                                {
                                    calculatedNextDate = lastDate.AddMonths((int)intervalMonths);
                                }
                            }
                        }
                    }

                    var newRow = masterWs.Row(++masterLastRow);
                    for (int col = 0; col < masterColCount; col++)
                    {
                        if (col == masterNextDateIdx && calculatedNextDate.HasValue)
                        {
                            var cell = newRow.Cell(col + 1);
                            cell.Value = calculatedNextDate.Value;
                            cell.Style.NumberFormat.Format = "yyyy-MM-dd";
                            continue;
                        }

                        if (col == 0)
                        {
                            newRow.Cell(1).Value = unitNumber;
                            continue;
                        }
                        int srcCol = headerMap[col];
                        if (srcCol >= 0)
                        {
                            var cell = ws.Cell(row, srcCol + 1);
                            newRow.Cell(col + 1).Value = cell.Value;
                        }
                        else
                        {
                            newRow.Cell(col + 1).Value = "";
                        }
                    }
                    row++;
                }
            }
            masterWb.SaveAs(outputFilePath);
        }
    }
} 