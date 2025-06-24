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

            // Normalize existing "Reoccurring" data in master file
            if (masterReoccurringIdx != -1)
            {
                var reoccurringColumn = masterWs.Column(masterReoccurringIdx + 1);
                // Start from row 3 to skip header and filter rows
                foreach (var cell in reoccurringColumn.CellsUsed(c => c.Address.RowNumber >= 3))
                {
                    if (cell.DataType == XLDataType.Text)
                    {
                        string originalValue = cell.GetString().Trim();
                        if (bool.TryParse(originalValue, out bool boolValue))
                        {
                            cell.Value = boolValue ? "True" : "False";
                        }
                    }
                }
            }

            int masterColCount = masterHeaders.Count;
            var lastRow = masterWs.LastRowUsed();
            int masterLastRow = lastRow != null ? lastRow.RowNumber() : 1; // 1 = header row, so data starts at 2

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

                string unitNumber = ExtractUnitNumberFromFileName(filePath);
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
                            if (reoccurringCell.DataType == XLDataType.Boolean) isReoccurring = reoccurringCell.GetBoolean();
                            else if (reoccurringCell.DataType == XLDataType.Text) bool.TryParse(reoccurringCell.GetString(), out isReoccurring);

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
                            var sourceCell = ws.Cell(row, srcCol + 1);
                            var targetCell = newRow.Cell(col + 1);

                            // For date columns, ensure proper conversion before copying
                            if (col == masterLastDateIdx || col == masterNextDateIdx)
                            {
                                if (TryParseDate(sourceCell, out DateTime dateValue))
                                {
                                    targetCell.Value = dateValue;
                                    targetCell.Style.NumberFormat.Format = "yyyy-MM-dd";
                                }
                                else
                                {
                                    targetCell.Value = sourceCell.Value;
                                    FormatDateCell(targetCell);
                                }
                            }
                            else
                            {
                                targetCell.Value = sourceCell.Value;
                                if (col == masterReoccurringIdx && targetCell.DataType == XLDataType.Text)
                                {
                                    string originalValue = targetCell.GetString().Trim();
                                    if (bool.TryParse(originalValue, out bool boolValue))
                                    {
                                        targetCell.Value = boolValue ? "True" : "False";
                                    }
                                }
                            }
                        }
                        else
                        {
                            newRow.Cell(col + 1).Value = "";
                        }
                    }
                    row++;
                }
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