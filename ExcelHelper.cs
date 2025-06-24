using System.Collections.Generic;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System;
using System.Linq;

namespace IMLoader
{
    public static class ExcelHelper
    {
        private static bool TryParseDate(IXLCell cell, out DateTime resultDate)
        {
            resultDate = default;
            try
            {
                // First try: If it's already a DateTime
                if (cell.DataType == XLDataType.DateTime)
                {
                    resultDate = cell.GetDateTime();
                    return true;
                }

                // Second try: If it's a number, try to convert from Excel serial date
                if (cell.DataType == XLDataType.Number && cell.TryGetValue(out double serialNumber))
                {
                    if (IsExcelDateSerial(serialNumber))
                    {
                        resultDate = DateTime.FromOADate(serialNumber);
                        return true;
                    }
                }

                // Third try: Parse as string with various formats
                string dateString = cell.GetString().Trim();
                if (!string.IsNullOrWhiteSpace(dateString))
                {
                    string[] formats = { "yyyy/MM/dd", "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd-MM-yyyy", "MM-dd-yyyy", "M/d/yyyy" };
                    if (DateTime.TryParseExact(dateString, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                    {
                        resultDate = parsedDate;
                        return true;
                    }
                    if (DateTime.TryParse(dateString, out DateTime parsedDate2))
                    {
                        resultDate = parsedDate2;
                        return true;
                    }
                }
            }
            catch { }
            return false;
        }

        private static bool IsExcelDateSerial(double number)
        {
            // Excel date serials typically fall between these values
            // 1 = January 1, 1900
            // 2958465 = December 31, 9999
            return number >= 1 && number <= 2958465;
        }

        private static void FormatDateCell(IXLCell cell)
        {
            if (cell == null) return;

            try
            {
                // If it's a number, try to convert from Excel serial date first
                if (cell.DataType == XLDataType.Number && cell.TryGetValue(out double serialNumber))
                {
                    if (IsExcelDateSerial(serialNumber))
                    {
                        try
                        {
                            var convertedDate = DateTime.FromOADate(serialNumber);
                            cell.Value = convertedDate;
                            cell.Style.NumberFormat.Format = "yyyy-MM-dd";
                            return;
                        }
                        catch { }
                    }
                }

                // If it's already a DateTime, just format it
                if (cell.DataType == XLDataType.DateTime)
                {
                    cell.Style.NumberFormat.Format = "yyyy-MM-dd";
                    return;
                }

                // For text values or any other type, try to parse as date
                if (TryParseDate(cell, out DateTime parsedDate))
                {
                    cell.Value = parsedDate;
                    cell.Style.NumberFormat.Format = "yyyy-MM-dd";
                }
            }
            catch (Exception)
            {
                // If any error occurs during conversion, try one last time to convert from serial number
                try
                {
                    if (cell.TryGetValue(out double lastChanceSerial) && IsExcelDateSerial(lastChanceSerial))
                    {
                        var finalDate = DateTime.FromOADate(lastChanceSerial);
                        cell.Value = finalDate;
                        cell.Style.NumberFormat.Format = "yyyy-MM-dd";
                    }
                }
                catch { }
            }
        }

        private static string NormalizeBooleanValue(IXLCell cell)
        {
            if (cell == null) return "False";

            if (cell.DataType == XLDataType.Boolean)
            {
                return cell.GetBoolean() ? "True" : "False";
            }
            
            if (cell.DataType == XLDataType.Text)
            {
                string value = cell.GetString().Trim();
                if (bool.TryParse(value, out bool boolValue))
                {
                    return boolValue ? "True" : "False";
                }
            }

            return "False";
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

            // Format existing date columns and normalize boolean values in master file
            if (masterReoccurringIdx != -1)
            {
                var reoccurringColumn = masterWs.Column(masterReoccurringIdx + 1);
                foreach (var cell in reoccurringColumn.CellsUsed())
                {
                    if (cell.Address.RowNumber > 2) // Skip header and filter rows
                    {
                        cell.Value = NormalizeBooleanValue(cell);
                    }
                }
            }

            // Format existing date columns in master file
            if (masterLastDateIdx != -1)
            {
                var lastDateColumn = masterWs.Column(masterLastDateIdx + 1);
                foreach (var cell in lastDateColumn.CellsUsed())
                {
                    if (cell.Address.RowNumber > 2) // Skip header and filter rows
                    {
                        FormatDateCell(cell);
                    }
                }
            }
            if (masterNextDateIdx != -1)
            {
                var nextDateColumn = masterWs.Column(masterNextDateIdx + 1);
                foreach (var cell in nextDateColumn.CellsUsed())
                {
                    if (cell.Address.RowNumber > 2) // Skip header and filter rows
                    {
                        FormatDateCell(cell);
                    }
                }
            }

            // Store all rows to be merged, along with their unit numbers
            var rowsToMerge = new List<(int unitNumber, List<XLCellValue> rowData)>();
            int masterColCount = masterHeaders.Count;

            // First, collect existing data from master file (after row 2)
            var existingRows = masterWs.Rows(3, masterWs.LastRowUsed()?.RowNumber() ?? 2);
            foreach (var row in existingRows)
            {
                var unitCell = row.Cell(1);
                if (!unitCell.IsEmpty())
                {
                    string unitStr = unitCell.GetString().Trim();
                    if (int.TryParse(unitStr, out int unitNum))
                    {
                        var rowData = new List<XLCellValue>();
                        for (int col = 1; col <= masterColCount; col++)
                        {
                            var cell = row.Cell(col);
                            if (col == masterReoccurringIdx)
                            {
                                rowData.Add(NormalizeBooleanValue(cell));
                            }
                            else
                            {
                                rowData.Add(cell.Value);
                            }
                        }
                        rowsToMerge.Add((unitNum, rowData));
                    }
                }
            }

            // Process files to merge
            foreach (var (filePath, sheetName) in filesToMerge)
            {
                using var wb = new XLWorkbook(filePath);
                if (!wb.Worksheets.TryGetWorksheet(sheetName, out var ws) || ws == null)
                    throw new ArgumentException($"Sheet '{sheetName}' not found in file '{filePath}'.");
                
                // Pre-format date columns in source file
                var headers = ws.Row(1).CellsUsed().Select(c => c.GetString()).ToList();
                var headerMap = new List<int>();
                for (int i = 0; i < masterHeaders.Count; i++)
                {
                    string masterHeader = masterHeaders[i].Trim();
                    int exactIdx = headers.FindIndex(h => string.Equals(h.Trim(), masterHeader, StringComparison.OrdinalIgnoreCase));
                    if (exactIdx >= 0)
                    {
                        headerMap.Add(exactIdx);
                        // Pre-format date columns in source file
                        if ((i == masterLastDateIdx || i == masterNextDateIdx) && exactIdx >= 0)
                        {
                            var dateColumn = ws.Column(exactIdx + 1);
                            foreach (var cell in dateColumn.CellsUsed())
                            {
                                if (cell.Address.RowNumber > 2)
                                {
                                    FormatDateCell(cell);
                                }
                            }
                        }
                    }
                    else
                    {
                        // Partial match logic remains the same...
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
                }

                string unitNumberStr = ExtractUnitNumberFromFileName(filePath);
                if (int.TryParse(unitNumberStr, out int unitNumber))
                {
                    int row = 2;
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
                                var lastDateCell = dataRow.Cell(srcLastDateIdx + 1);

                                // Ensure Last Date is formatted before using it
                                FormatDateCell(lastDateCell);

                                bool isReoccurring = false;
                                if (reoccurringCell.DataType == XLDataType.Boolean)
                                {
                                    isReoccurring = reoccurringCell.GetBoolean();
                                }
                                else if (reoccurringCell.DataType == XLDataType.Text)
                                {
                                    bool.TryParse(reoccurringCell.GetString().Trim(), out isReoccurring);
                                }

                                if (!isReoccurring && (nextDateCell == null || nextDateCell.IsEmpty()))
                                {
                                    var intervalCell = dataRow.Cell(srcIntervalIdx + 1);
                                    if (TryParseDate(lastDateCell, out DateTime lastDate) && intervalCell.TryGetValue(out double intervalMonths))
                                    {
                                        calculatedNextDate = lastDate.AddMonths((int)intervalMonths);
                                    }
                                }
                            }
                        }

                        var rowData = new List<XLCellValue>();
                        
                        // Add unit number as first column
                        rowData.Add(unitNumber.ToString());

                        // Add rest of the columns
                        for (int col = 1; col < masterColCount; col++)
                        {
                            if (col == masterNextDateIdx && calculatedNextDate.HasValue)
                            {
                                rowData.Add(calculatedNextDate.Value);
                                continue;
                            }

                            int srcCol = headerMap[col];
                            if (srcCol >= 0)
                            {
                                var sourceCell = ws.Cell(row, srcCol + 1);
                                
                                // Handle date columns
                                if (col == masterLastDateIdx || col == masterNextDateIdx)
                                {
                                    if (TryParseDate(sourceCell, out DateTime dateValue))
                                    {
                                        rowData.Add(dateValue);
                                    }
                                    else
                                    {
                                        rowData.Add(sourceCell.Value);
                                    }
                                }
                                // Handle Reoccurring column
                                else if (col == masterReoccurringIdx)
                                {
                                    rowData.Add(NormalizeBooleanValue(sourceCell));
                                }
                                else
                                {
                                    rowData.Add(sourceCell.Value);
                                }
                            }
                            else
                            {
                                rowData.Add("");
                            }
                        }

                        rowsToMerge.Add((unitNumber, rowData));
                        row++;
                    }
                }
            }

            // Sort all rows by unit number
            rowsToMerge.Sort((a, b) => a.unitNumber.CompareTo(b.unitNumber));

            // Clear existing data (after row 2) and write sorted data
            var lastUsedRow = masterWs.LastRowUsed();
            if (lastUsedRow != null && lastUsedRow.RowNumber() > 2)
            {
                masterWs.Rows(3, lastUsedRow.RowNumber()).Delete();
            }

            // Write sorted data back to worksheet
            int currentRow = 3;
            foreach (var (_, rowData) in rowsToMerge)
            {
                var newRow = masterWs.Row(currentRow);
                for (int col = 0; col < rowData.Count; col++)
                {
                    var cell = newRow.Cell(col + 1);
                    cell.Value = rowData[col];

                    // Apply formatting for date columns
                    if (col == masterLastDateIdx || col == masterNextDateIdx)
                    {
                        FormatDateCell(cell);
                    }
                }
                currentRow++;
            }

            // Final pass to ensure all date columns are properly formatted
            if (masterLastDateIdx != -1)
            {
                foreach (var cell in masterWs.Column(masterLastDateIdx + 1).CellsUsed())
                {
                    if (cell.Address.RowNumber > 2)
                    {
                        FormatDateCell(cell);
                    }
                }
            }
            if (masterNextDateIdx != -1)
            {
                foreach (var cell in masterWs.Column(masterNextDateIdx + 1).CellsUsed())
                {
                    if (cell.Address.RowNumber > 2)
                    {
                        FormatDateCell(cell);
                    }
                }
            }

            masterWb.SaveAs(outputFilePath);
        }
    }
} 