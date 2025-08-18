using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace ExcelDiagnostics.Infrastructure
{
    // 1. Define core interfaces and classes
    public interface IExcelItem
    {
        string RowError { get; }
        void SetError(string error);
        bool IsValid();
    }

    public class ExcelItem<T>
    {
        public int? Index { get; set; }
        public T Value { get; set; }
        public T DefaultValue { get; set; }
        public Func<string, (T parsedValue, bool isValid)> CustomParse { get; set; }
        public string Error { get; private set; }

        public void SetError(string error) => Error = error;
        public void ClearError() => Error = null;
    }

    public class ExcelReadResult<T> where T : IExcelItem
    {
        public List<T> Items { get; } = new List<T>();
        public int TotalRows => Items.Count;
        public int ValidRows => Items.Count(row => row.IsValid());
        public int ErrorCount => Items.Sum(row => row.IsValid() ? 0 : 1);

        public bool IsValid() => Items.All(row => row.IsValid());

        public void GenerateErrorsExcelFile(string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Errors");
                var properties = typeof(T).GetProperties()
                    .Where(p => p.PropertyType.IsGenericType &&
                                p.PropertyType.GetGenericTypeDefinition() == typeof(ExcelItem<>))
                    .ToList();

                // Create headers
                for (int i = 0; i < properties.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = properties[i].Name;
                }
                worksheet.Cells[1, properties.Count + 1].Value = "RowError";

                // Populate data
                for (int rowIdx = 0; rowIdx < Items.Count; rowIdx++)
                {
                    var row = Items[rowIdx];
                    var excelRow = rowIdx + 2; // Start after header

                    for (int colIdx = 0; colIdx < properties.Count; colIdx++)
                    {
                        var prop = properties[colIdx];
                        var excelItem = prop.GetValue(row);
                        var valueProp = excelItem.GetType().GetProperty("Value");
                        var errorProp = excelItem.GetType().GetProperty("Error");

                        var cell = worksheet.Cells[excelRow, colIdx + 1];
                        cell.Value = valueProp.GetValue(excelItem) ?? "";

                        // Apply error formatting
                        if (errorProp.GetValue(excelItem) is string error && !string.IsNullOrEmpty(error))
                        {
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCoral);
                            cell.AddComment(error, "Error");
                        }
                    }

                    // Add row-level error
                    if (!string.IsNullOrEmpty(row.RowError))
                    {
                        var errorCell = worksheet.Cells[excelRow, properties.Count + 1];
                        errorCell.Value = row.RowError;
                        errorCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        errorCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
                    }
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }
    }

    // 2. Implement the diagnostic helper
    public class ExcelDiagnosticHelper
    {
        public ExcelDiagnosticHelper()
        {
            // Ensure license is configured before any Excel operations
            ExcelLicenseConfig.ConfigureLicense();
        }
        public ExcelReadResult<T> ReadExcel<T>(Stream fileStream) where T : IExcelItem, new()
        {
            var result = new ExcelReadResult<T>();

            using (var package = new ExcelPackage(fileStream))
            {

                var worksheet = package.Workbook.Worksheets[0];
                var properties = typeof(T).GetProperties()
                    .Where(p => p.PropertyType.IsGenericType &&
                                p.PropertyType.GetGenericTypeDefinition() == typeof(ExcelItem<>))
                    .ToList();

                // Create column mapping
                var columnMap = new Dictionary<int, PropertyInfo>();
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var header = worksheet.Cells[1, col].Text.Trim();
                    var prop = properties.FirstOrDefault(p =>
                        p.Name.Equals(header, StringComparison.OrdinalIgnoreCase));

                    if (prop != null) columnMap.Add(col, prop);
                }

                // Process rows
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var item = new T();

                    foreach (var mapping in columnMap)
                    {
                        int col = mapping.Key;
                        var prop = mapping.Value;
                        var cellValue = worksheet.Cells[row, col].Text;
                        var excelItem = prop.GetValue(item);
                        var itemType = prop.PropertyType.GetGenericArguments()[0];

                        // Use reflection to access ExcelItem properties
                        var indexProp = prop.PropertyType.GetProperty("Index");
                        var valueProp = prop.PropertyType.GetProperty("Value");
                        var defaultValueProp = prop.PropertyType.GetProperty("DefaultValue");
                        var customParseProp = prop.PropertyType.GetProperty("CustomParse");
                        var setErrorMethod = prop.PropertyType.GetMethod("SetError");

                        // Try parsing value
                        try
                        {
                            if (string.IsNullOrWhiteSpace(cellValue))
                            {
                                var defaultValue = defaultValueProp.GetValue(excelItem);
                                if (defaultValue != null)
                                {
                                    valueProp.SetValue(excelItem, defaultValue);
                                }
                                else
                                {
                                    setErrorMethod.Invoke(excelItem, new object[] { "Mandatory field missing" });
                                }
                            }
                            else
                            {
                                var customParse = customParseProp.GetValue(excelItem) as Delegate;
                                if (customParse != null)
                                {
                                    var parseResult = customParse.DynamicInvoke(cellValue);
                                    var parsedValue = parseResult.GetType().GetField("parsedValue").GetValue(parseResult);
                                    var isValid = (bool)parseResult.GetType().GetField("isValid").GetValue(parseResult);

                                    if (isValid)
                                    {
                                        valueProp.SetValue(excelItem, parsedValue);
                                    }
                                    else
                                    {
                                        setErrorMethod.Invoke(excelItem, new object[] { "Custom validation failed" });
                                    }
                                }
                                else
                                {
                                    // Built-in type conversion
                                    var convertedValue = Convert.ChangeType(cellValue, itemType);
                                    valueProp.SetValue(excelItem, convertedValue);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            setErrorMethod.Invoke(excelItem, new object[] { $"Invalid format: {ex.Message}" });
                        }
                    }

                    result.Items.Add(item);
                }
            }
            return result;
        }
        public static class ExcelLicenseConfig
        {
            private static bool _licenseSet = false;
            private static readonly object _lock = new object();

            public static void ConfigureLicense()
            {
                if (!_licenseSet)
                {
                    lock (_lock)
                    {
                        if (!_licenseSet)
                        {
                            // For non-commercial use:
                            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                            // For commercial use (uncomment and add your key):
                            // ExcelPackage.LicenseContext = LicenseContext.Commercial;
                            // ExcelPackage.LicenseKey = "your-commercial-license-key";

                            _licenseSet = true;
                        }
                    }
                }
            }
        }
    }

    // 3. Example model implementation
    public class ExcelInvoiceResult : IExcelItem
    {
        public string RowError { get; private set; }

        public void SetError(string error) => RowError = error;

        public bool IsValid() => string.IsNullOrEmpty(RowError) &&
            ItemID.Error == null &&
            ItemCode.Error == null &&
            ItemCodeSup.Error == null &&
            ItemCodeCust.Error == null;

        // Excel items with custom configurations
        public ExcelItem<int> ItemID { get; set; } = new ExcelItem<int>
        {
            Index = 1,
            DefaultValue = 0
        };

        public ExcelItem<string> ItemCode { get; set; } = new ExcelItem<string>
        {
            Index = 2,
            CustomParse = input =>
            {
                bool isValid = !string.IsNullOrEmpty(input) && input.StartsWith("ITM-");
                return (input, isValid);
            }
        };

        public ExcelItem<string> ItemCodeSup { get; set; } = new ExcelItem<string>
        {
            Index = 3
        };

        public ExcelItem<int> ItemCodeCust { get; set; } = new ExcelItem<int>
        {
            Index = 4,
            CustomParse = input =>
            {
                if (int.TryParse(input, out int value) && value > 0)
                    return (value, true);
                return (0, false);
            }
        };
    }

    // 4. Usage example in a controller
    ////public class ExcelUploadController
    ////{
    ////    private readonly ExcelDiagnosticHelper _excelHelper = new();

    ////    public IActionResult UploadInvoice(IFormFile file)
    ////    {
    ////        using var stream = file.OpenReadStream();
    ////        var result = _excelHelper.ReadExcel<ExcelInvoiceResult>(stream);

    ////        if (!result.IsValid())
    ////        {
    ////            var errorPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}_errors.xlsx");
    ////            result.GenerateErrorsExcelFile(errorPath);
    ////            return File(System.IO.File.ReadAllBytes(errorPath),
    ////                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ////                      "InvoiceErrors.xlsx");
    ////        }

    ////        // Process valid data
    ////        foreach (var item in result.Items.Where(i => i.IsValid()))
    ////        {
    ////            // Access typed values:
    ////            int id = item.ItemID.Value;
    ////            string code = item.ItemCode.Value;
    ////            // ... further processing
    ////        }

    ////        return Ok("Processing completed successfully");
    ////    }
    ////}
}
