using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Models
{
    /// <summary>
    /// Final diagnostic result 
    /// </summary>
    public class ExcelDiagnosticResult<TModel>
    {
        public List<ExcelRowResult<TModel>> Rows { get; } = new();
        public bool IsValid() => Rows.All(r => r.IsValid());

        // Diagnostics summary (counters) - i will push more details later as you demand .
        public int TotalRows => Rows.Count;
        public int ValidRowsCount => Rows.Count(r => r.IsValid());
        public int InvalidRowsCount => Rows.Count - ValidRowsCount;
        public int TotalCellErrorsCount => Rows.Sum(r => r.CellErrors.Count);
        public int TotalRowErrorsCount => Rows.Sum(r => r.RowErrors.Count);
        public int TotalCellWarningsCount => Rows.Sum(r => r.CellWarnings.Values.Sum(w => w.Count));
        public int TotalRowWarningsCount => Rows.Sum(r => r.RowWarnings.Count);


        /// <summary>
        /// Generate an Excel file with three sheets: DataWithErrors, Warnings, and Errors
        /// ,Hover on cells to see error/warning details.
        /// </summary>
        public void GenerateErrorsExcelFile(string outputPath, string? originalWorkbookPath = null)
        {
            using var outPackage = new ExcelPackage();

            if (!string.IsNullOrWhiteSpace(originalWorkbookPath) && File.Exists(originalWorkbookPath))
            {
                try
                {
                    using var inPackage = new ExcelPackage(new FileInfo(originalWorkbookPath));
                    var source = inPackage.Workbook.Worksheets.FirstOrDefault();
                    if (source != null)
                    {
                        var copy = outPackage.Workbook.Worksheets.Add("DataWithErrors");

                        CopyWorksheetData(source, copy);

                        HighlightCells(copy);
                    }
                }
                catch (Exception ex)
                {
                    var copy = outPackage.Workbook.Worksheets.Add("DataWithErrors");
                    copy.Cells[1, 1].Value = $"Error copying original data: {ex.Message}";
                }
            }

            // 2)  Errors 
            CreateErrorsSheet(outPackage);

            // 3)  Warnings sheet
            CreateWarningsSheet(outPackage);

            var fi = new FileInfo(outputPath);
            if (fi.Exists) fi.Delete();
            outPackage.SaveAs(fi);
        }

        private void CopyWorksheetData(ExcelWorksheet source, ExcelWorksheet target)
        {
            var startRow = source.Dimension?.Start.Row ?? 1;
            var endRow = source.Dimension?.End.Row ?? 1;
            var startCol = source.Dimension?.Start.Column ?? 1;
            var endCol = source.Dimension?.End.Column ?? 1;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var sourceCell = source.Cells[row, col];
                    var targetCell = target.Cells[row, col];

                    if (sourceCell.Value != null)
                    {
                        targetCell.Value = sourceCell.Value;

                        if (sourceCell.Style.Font.Bold)
                            targetCell.Style.Font.Bold = true;
                    }
                }
            }
        }

        private void HighlightCells(ExcelWorksheet worksheet)
        {
            foreach (var row in Rows)
            {
                foreach (var ce in row.CellErrors)
                {
                    HighlightCell(worksheet, row, ce.Key, System.Drawing.Color.LightPink, $"ERROR: {ce.Value}");
                }

                foreach (var cw in row.CellWarnings)
                {
                    foreach (var warning in cw.Value)
                    {
                        HighlightCell(worksheet, row, cw.Key, System.Drawing.Color.LightYellow, $"WARNING: {warning}");
                    }
                }
            }
        }

        private void HighlightCell(ExcelWorksheet worksheet, ExcelRowResult<TModel> row, string propertyName, System.Drawing.Color color, string comment)
        {
            var prop = row.Data!.GetType().GetProperty(propertyName);
            if (prop?.GetValue(row.Data) is object excelItemObj)
            {
                dynamic excelItem = excelItemObj;
                int col = (int)(excelItem.Index ?? -1);

                if (col > 0 && row.SourceRowNumber > 0)
                {
                    var cell = worksheet.Cells[row.SourceRowNumber, col];
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(color);

                    
                    if (cell.Comment == null)
                    {
                        cell.AddComment(comment, "Validator");
                    }
                    
                }
            }
        }

        private void CreateWarningsSheet(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Warnings");

            // Headers
            ws.Cells[1, 1].Value = "Row Number";
            ws.Cells[1, 2].Value = "Column";
            ws.Cells[1, 3].Value = "Cell";
            ws.Cells[1, 4].Value = "Warning Description";

            int currentRow = 2;

            foreach (var row in Rows)
            {
                foreach (var cw in row.CellWarnings)
                {
                    foreach (var warning in cw.Value)
                    {
                        string cellAddr = GetCellAddress(row, cw.Key);
                        ws.Cells[currentRow, 1].Value = row.SourceRowNumber;
                        ws.Cells[currentRow, 2].Value = cw.Key;
                        ws.Cells[currentRow, 3].Value = cellAddr;
                        ws.Cells[currentRow, 4].Value = warning;
                        currentRow++;
                    }
                }

                foreach (var warning in row.RowWarnings)
                {
                    ws.Cells[currentRow, 1].Value = row.SourceRowNumber;
                    ws.Cells[currentRow, 2].Value = "(Row)";
                    ws.Cells[currentRow, 3].Value = "-";
                    ws.Cells[currentRow, 4].Value = warning;
                    currentRow++;
                }
            }

            FormatSheet(ws, 4);
        }

        private void CreateErrorsSheet(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Errors");

            // Headers
            ws.Cells[1, 1].Value = "Row Number";
            ws.Cells[1, 2].Value = "Column";
            ws.Cells[1, 3].Value = "Cell";
            ws.Cells[1, 4].Value = "Error Description";

            int currentRow = 2;

            foreach (var row in Rows)
            {
                foreach (var ce in row.CellErrors)
                {
                    string cellAddr = GetCellAddress(row, ce.Key);
                    ws.Cells[currentRow, 1].Value = row.SourceRowNumber;
                    ws.Cells[currentRow, 2].Value = ce.Key;
                    ws.Cells[currentRow, 3].Value = cellAddr;
                    ws.Cells[currentRow, 4].Value = ce.Value;
                    currentRow++;
                }

                foreach (var error in row.RowErrors)
                {
                    ws.Cells[currentRow, 1].Value = row.SourceRowNumber;
                    ws.Cells[currentRow, 2].Value = "(Row)";
                    ws.Cells[currentRow, 3].Value = "-";
                    ws.Cells[currentRow, 4].Value = error;
                    currentRow++;
                }
            }

            FormatSheet(ws, 4);
        }

        private string GetCellAddress(ExcelRowResult<TModel> row, string propertyName)
        {
           
            var prop = row.Data!.GetType().GetProperty(propertyName);
            if (prop?.GetValue(row.Data) is object excelItemObj)
            {
                dynamic excelItem = excelItemObj;
                int col = (int)(excelItem.Index ?? -1);
                if (col > 0 && row.SourceRowNumber > 0)
                {
                    return new ExcelCellAddress(row.SourceRowNumber, col).Address;
                }
            }
            
            return "-";
        }

        private void FormatSheet(ExcelWorksheet ws, int columnCount)
        {
           
            // Bold H
            using (var header = ws.Cells[1, 1, 1, columnCount])
            {
                header.Style.Font.Bold = true;
                header.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                header.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            ws.Cells.AutoFitColumns();
           
        }

    }
}
