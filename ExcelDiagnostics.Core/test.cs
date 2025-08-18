using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace ExcelDiagnosticHelper
{
    #region Core Contracts

    /// <summary>
    /// Optional contract for row-level validators.
    /// Implement Validate to add row errors or cell errors (via row.SetError/SetCellError).
    /// </summary>
    public interface IRowValidator<TModel>
    {
        void Validate(ExcelRowResult<TModel> row);
    }

    /// <summary>
    /// Contract for reading Excel into typed models with validation.
    /// </summary>
    public interface IExcelReader
    {
        ExcelDiagnosticResult<TModel> ReadExcel<TModel>(
            string filePath,
            int headerRow = 1,
            IEnumerable<IRowValidator<TModel>>? rowValidators = null) where TModel : new();

        ExcelDiagnosticResult<TModel> ReadExcel<TModel>(
            Stream excelStream,
            int headerRow = 1,
            IEnumerable<IRowValidator<TModel>>? rowValidators = null) where TModel : new();
    }

    #endregion

    #region Cell Wrapper & Row/Result Containers

    /// <summary>
    /// Wrapper for a cell. Holds column index, parsed value, default value,
    /// parse function, validators, and error message. Provides IsValid/SetError.
    /// </summary>
    public class ExcelItem<T>
    {
        /// <summary> 1-based column index in the worksheet. </summary>
        public int? Index { get; set; }

        /// <summary> Parsed value. </summary>
        public T? Value { get; set; }

        /// <summary> Default value used when the cell is empty or parsing fails (if desired). </summary>
        public T? DefaultValue { get; set; }

        /// <summary> Optional custom parse function: input raw => parsed T. If null, default conversion is used. </summary>
        public Func<object?, T?>? CustomParse { get; set; }

        /// <summary> Optional custom validators that must all pass. </summary>
        public List<Func<T?, bool>> Validators { get; } = new();

        /// <summary> Error message for this cell (null/empty if valid). </summary>
        public string? Error { get; private set; }

        public bool IsValid() => string.IsNullOrWhiteSpace(Error);

        public void SetError(string error) => Error = error;

        /// <summary>
        /// Set true if this cell must have a value (non-null / non-empty for strings)
        /// </summary>
        public bool Required { get; set; } = false;

        /// <summary>
        /// Utility to add a validator with a message. If validator returns false, sets the message as error.
        /// </summary>
        public void AddValidator(Func<T?, bool> predicate, string errorMessage)
        {
            Validators.Add(value =>
            {
                var ok = predicate(value);
                if (!ok && string.IsNullOrWhiteSpace(Error))
                    SetError(errorMessage);
                return ok;
            });
        }



        /// <summary>
        /// Run all validators including the required check
        /// </summary>
        public void RunValidators()
        {
            // First check Required
            if (Required)
            {
                if (Value == null || (Value is string s && string.IsNullOrWhiteSpace(s)))
                {
                    SetError("Value is required.");
                    return; // Stop if required fails
                }
            }

            // Run custom validators
            foreach (var validator in Validators)
            {
                bool ok = validator(Value);
                if (!ok) break; // Stop at first failed validator
            }
        }
    }


    /// <summary>
    /// Holds one imported row: the typed Data, per-cell errors, row errors, and the source row number.
    /// </summary>
    public class ExcelRowResult<TModel>
    {
        public int SourceRowNumber { get; }
        public TModel Data { get; }

        // --- Existing diagnostics ---
        public Dictionary<string, string> CellErrors { get; } = new(); // PropertyName -> Error
        public List<string> RowErrors { get; } = new();

        // --- NEW: Support for warnings (non-fatal issues) ---
        public Dictionary<string, List<string>> CellWarnings { get; } = new(); // PropertyName -> warnings
        public List<string> RowWarnings { get; } = new();

        public ExcelRowResult(int sourceRowNumber, TModel data)
        {
            SourceRowNumber = sourceRowNumber;
            Data = data;
        }

        // A row is valid if no errors (warnings are ignored for validity check)
        public bool IsValid() => !CellErrors.Any() && !RowErrors.Any();

        public void SetError(string error) => RowErrors.Add(error);
        public void SetCellError(string propertyName, string error)
            => CellErrors[propertyName] = error;

        // --- NEW: Helpers for warnings ---
        public void SetRowWarning(string warning) => RowWarnings.Add(warning);
        public void SetCellWarning(string propertyName, string warning)
        {
            if (!CellWarnings.TryGetValue(propertyName, out var list))
            {
                list = new List<string>();
                CellWarnings[propertyName] = list;
            }
            list.Add(warning);
        }
    }

    /// <summary>
    /// Final diagnostic result across all rows. Includes generation of an Errors workbook.
    /// </summary>
    public class ExcelDiagnosticResult<TModel>
    {
        public List<ExcelRowResult<TModel>> Rows { get; } = new();

        public bool IsValid() => Rows.All(r => r.IsValid());

        // --- NEW: Diagnostics summary (counts) ---

        /// <summary>Total number of rows parsed (excluding empty ones).</summary>
        public int TotalRows => Rows.Count;

        /// <summary>Number of rows with no errors.</summary>
        public int ValidRowsCount => Rows.Count(r => r.IsValid());

        /// <summary>Number of rows that have at least one error.</summary>
        public int InvalidRowsCount => Rows.Count - ValidRowsCount;

        /// <summary>Total number of individual cell errors across all rows.</summary>
        public int TotalCellErrorsCount => Rows.Sum(r => r.CellErrors.Count);

        /// <summary>Total number of row-level errors across all rows.</summary>
        public int TotalRowErrorsCount => Rows.Sum(r => r.RowErrors.Count);

        // --- NEW: Warnings counts ---
        public int TotalCellWarningsCount => Rows.Sum(r => r.CellWarnings.Values.Sum(w => w.Count));
        public int TotalRowWarningsCount => Rows.Sum(r => r.RowWarnings.Count);

        /// <summary>
        /// Generates an Excel workbook with a single "Errors" sheet listing all errors.
        /// Saves to the specified path (overwrites if exists).
        /// </summary>
        public void GenerateErrorsExcelFile(string path)
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Errors");

            // Headers
            ws.Cells[1, 1].Value = "Row Number";
            ws.Cells[1, 2].Value = "Column";
            ws.Cells[1, 3].Value = "Cell";
            ws.Cells[1, 4].Value = "Error Description";

            int r = 2;

            foreach (var row in Rows)
            {
                // Cell-level errors
                foreach (var ce in row.CellErrors)
                {
                    var prop = row.Data!.GetType().GetProperty(ce.Key);
                    int columnIndex = -1;
                    if (prop != null && prop.GetValue(row.Data) is object excelItemObj)
                    {
                        dynamic excelItemDyn = excelItemObj;
                        try { columnIndex = (int)(excelItemDyn.Index ?? -1); }
                        catch { columnIndex = -1; }
                    }

                    string cellAddr = columnIndex > 0
                        ? new ExcelCellAddress(row.SourceRowNumber, columnIndex).Address
                        : "-";

                    ws.Cells[r, 1].Value = row.SourceRowNumber; // Row number in source Excel
                    ws.Cells[r, 2].Value = ce.Key;              // Property name
                    ws.Cells[r, 3].Value = cellAddr;            // Cell address
                    ws.Cells[r, 4].Value = ce.Value;            // Error description
                    r++;
                }

                // Row-level errors
                foreach (var re in row.RowErrors)
                {
                    ws.Cells[r, 1].Value = row.SourceRowNumber;
                    ws.Cells[r, 2].Value = "(Row)";
                    ws.Cells[r, 3].Value = "-";
                    ws.Cells[r, 4].Value = re;
                    r++;
                }
            }

            // Format headers
            using (var header = ws.Cells[1, 1, 1, 4])
            {
                header.Style.Font.Bold = true;
            }
            ws.Cells.AutoFitColumns();

            var fi = new FileInfo(path);
            if (fi.Exists) fi.Delete();
            package.SaveAs(fi);
        }
    }


    #endregion

    #region Reader Implementation 

    public class ExcelReader : IExcelReader
    {
        public ExcelDiagnosticResult<TModel> ReadExcel<TModel>(
            string filePath,
            int headerRow = 1,
            IEnumerable<IRowValidator<TModel>>? rowValidators = null) where TModel : new()
        {
            using var fs = File.OpenRead(filePath);
            return ReadExcel<TModel>(fs, headerRow, rowValidators);
        }

        public ExcelDiagnosticResult<TModel> ReadExcel<TModel>(
            Stream excelStream,
            int headerRow = 1,
            IEnumerable<IRowValidator<TModel>>? rowValidators = null) where TModel : new()
        {
            if (!excelStream.CanSeek)
            {
                var temp = new MemoryStream();
                excelStream.CopyTo(temp);
                temp.Position = 0;
                excelStream = temp;
            }

            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using var package = new ExcelPackage(excelStream);
            var ws = package.Workbook.Worksheets.FirstOrDefault()
                     ?? throw new InvalidOperationException("Workbook does not contain any worksheets.");

            // --- 1. Reflection: obtain ExcelItem<T> properties once ---
            var props = GetExcelItemProperties<TModel>();
            var mapping = BuildColumnMapping<TModel>(props);

            var dim = ws.Dimension
                      ?? throw new InvalidOperationException("Worksheet is empty.");
            int startRow = Math.Max(1, headerRow + 1);
            int endRow = dim.End.Row;

            var result = new ExcelDiagnosticResult<TModel>();

            // --- 2. Iterate rows ---
            for (int row = startRow; row <= endRow; row++)
            {
                if (IsRowEmpty(ws, row, mapping.Values.Select(m => m.ColumnIndex))) continue;

                var model = new TModel();
                var rowResult = new ExcelRowResult<TModel>(row, model);

                // --- 3. Per-cell: parse value ---
                foreach (var (prop, meta) in mapping)
                {
                    dynamic? item = prop.GetValue(model);
                    if (item == null)
                    {
                        item = Activator.CreateInstance(prop.PropertyType);
                        prop.SetValue(model, item);
                    }

                    int col = meta.ColumnIndex;
                    item.Index = col;

                    object? raw = ws.Cells[row, col].Value;

                    try
                    {
                        if (IsNullOrEmptyLike(raw))
                            item.Value = item.DefaultValue;
                        else if (item.CustomParse is not null)
                            item.Value = item.CustomParse(raw);
                        else
                            item.Value = ConvertToGeneric(item, raw, prop.PropertyType);
                    }
                    catch (Exception ex)
                    {
                        item.Value = item.DefaultValue;
                        item.SetError($"Invalid '{prop.Name}' value at row {row}, col {col}: {ex.Message}");
                    }
                }

                // --- 4. Run row-level validators (can add cell validators dynamically) ---
                if (rowValidators != null)
                {
                    foreach (var validator in rowValidators)
                    {
                        try
                        {
                            validator.Validate(rowResult);
                        }
                        catch (Exception ex)
                        {
                            rowResult.SetError($"Row validator threw an exception: {ex.Message}");
                        }
                    }
                }

                // --- 5. Run cell-level validators (including those added by row validators) ---
                foreach (var (prop, _) in mapping)
                {
                    dynamic? item = prop.GetValue(model);
                    if (item == null) continue;

                    // Skip if parsing already failed
                    if (!item.IsValid())
                    {
                        rowResult.SetCellError(prop.Name, item.Error!);
                        continue;
                    }

                    // Run validators safely using ExcelItem<T>.RunValidators()
                    item.RunValidators();

                    if (!item.IsValid())
                        rowResult.SetCellError(prop.Name, item.Error!);
                }

                // --- 6. Add row to result ---
                result.Rows.Add(rowResult);
            }

            return result;
        }


        #region Helpers

        private static bool IsRowEmpty(ExcelWorksheet ws, int row, IEnumerable<int> columns)
        {
            foreach (var c in columns)
            {
                var val = ws.Cells[row, c].Value;
                if (!IsNullOrEmptyLike(val)) return false;
            }
            return true;
        }

        private static bool IsNullOrEmptyLike(object? value)
        {
            if (value is null) return true;
            if (value is string s) return string.IsNullOrWhiteSpace(s);
            return false;
        }

        private static IEnumerable<PropertyInfo> GetExcelItemProperties<TModel>()
            => typeof(TModel)
               .GetProperties(BindingFlags.Public | BindingFlags.Instance)
               .Where(p => p.PropertyType.IsGenericType &&
                           p.PropertyType.GetGenericTypeDefinition() == typeof(ExcelItem<>))
               .OrderBy(OrderByDeclarationOrder);

        private static int OrderByDeclarationOrder(PropertyInfo p) => p.MetadataToken;

        private sealed class ColumnMeta
        {
            public int ColumnIndex { get; init; }
        }

        private static Dictionary<PropertyInfo, ColumnMeta> BuildColumnMapping<TModel>(IEnumerable<PropertyInfo> props)
        {
            var mapping = new Dictionary<PropertyInfo, ColumnMeta>();
            int running = 1;
            foreach (var p in props)
            {
                dynamic? item = Activator.CreateInstance(p.PropertyType);
                int colIndex = item?.Index ?? running++;
                mapping[p] = new ColumnMeta { ColumnIndex = colIndex };
            }
            return mapping;
        }

        private static object? ConvertToGeneric(dynamic excelItem, object? raw, Type excelItemType)
        {
            Type t = excelItemType.GetGenericArguments()[0];
            if (raw is null) return null;

            if (t == typeof(string))
                return raw.ToString()?.Trim();

            var underlying = Nullable.GetUnderlyingType(t);
            if (underlying != null)
            {
                if (IsNullOrEmptyLike(raw)) return null;
                return ChangeType(raw, underlying);
            }

            return ChangeType(raw, t);
        }

        private static object? ChangeType(object value, Type targetType)
        {
            try
            {
                if (targetType == typeof(Guid))
                {
                    if (value is Guid g) return g;
                    return Guid.Parse(value.ToString()!);
                }

                if (targetType.IsEnum)
                {
                    if (value is string es) return Enum.Parse(targetType, es, ignoreCase: true);
                    return Enum.ToObject(targetType, Convert.ChangeType(value, Enum.GetUnderlyingType(targetType)));
                }

                var converter = TypeDescriptor.GetConverter(targetType);
                if (converter.CanConvertFrom(value.GetType()))
                    return converter.ConvertFrom(value);

                return Convert.ChangeType(value, targetType);
            }
            catch
            {
                throw new InvalidCastException($"Cannot convert '{value}' to {targetType.Name}");
            }
        }

        #endregion
    }

    #endregion

    #region Sample Model & Validators

    public class ExcelInvoiceRow
    {
        public ExcelItem<int> OrderId { get; set; } = new() { Index = 1, Required = true };
        public ExcelItem<string> ItemCode { get; set; } = new() { Index = 2,
            Required = true ,
            CustomParse = x=> x.ToString()?.Trim()
        };
        public ExcelItem<decimal?> Amount { get; set; } = new() { Index = 3, Required = true };
        public ExcelItem<string> Type { get; set; } = new() { Index = 4, Required = true };
        public ExcelItem<string> CustomerName { get; set; } = new() { Index = 5, Required = true };
    }

    public class SampleInvoiceRowValidator : IRowValidator<ExcelInvoiceRow>
    {
        public void Validate(ExcelRowResult<ExcelInvoiceRow> row)
        {
            var data = row.Data;

            //// Example validators
            data.ItemCode.AddValidator(v => !string.IsNullOrWhiteSpace(v) && v!.Length >= 3,
                "ItemCode must have at least 3 characters."); 
            data.ItemCode.AddValidator(v => !string.IsNullOrWhiteSpace(v) && v!.Length >= 3,
                "ItemCode must have at least 3 characters.");

            data.Amount.AddValidator(v => v.HasValue, "Amount must be provided.");

            data.OrderId.AddValidator(v => v >=5 , "OrderId above or equal 5");



            //// for Row-level rules
            var type = (data.Type.Value ?? "Sale").Trim();
            var amount = data.Amount.Value;

            if (string.Equals(type, "Refund", StringComparison.OrdinalIgnoreCase))
            {
                if (amount <= 0)
                    row.SetError("For 'Refund' rows, Amount must be positive.");
            }
            
        }
    }

    #endregion
}
