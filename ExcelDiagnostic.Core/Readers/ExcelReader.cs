using ExcelDataReader;
using ExcelDiagnostic.Core.Contracts;
using ExcelDiagnostic.Core.Models;
using ExcelDiagnostic.Core.Mapping;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelDiagnostic.Core.Utils;

namespace ExcelDiagnostic.Core.Readers
{
    /// <summary>
    /// Single concrete reader that uses ExcelDataReader could handle (.xls & .xlsx).
    /// warning :for now it reads first sheet of excel , but later i will add support for multiple sheets.
    /// </summary>
    public class ExcelReader : IExcelReader
    {
        static ExcelReader()// For static constructor to register CodePagesEncodingProvider xls files support
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        //add this you can pass only path of excel file;
        public ExcelDiagnosticResult<TModel> ReadExcel<TModel>(string filePath, int? headerRow = 1, IEnumerable<IRowValidator<TModel>>? rowValidators = null) where TModel : new()
        {
            using var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return ReadExcel<TModel>(fs, headerRow, rowValidators);
        }

        public ExcelDiagnosticResult<TModel> ReadExcel<TModel>(Stream excelStream, int? headerRow = 1, IEnumerable<IRowValidator<TModel>>? rowValidators = null) where TModel : new()
        {
            var result = new ExcelDiagnosticResult<TModel>();
            foreach (var row in ReadExcelStream<TModel>(excelStream, headerRow, rowValidators))
            {
                result.Rows.Add(row);
            }
            return result;
        }

        public IEnumerable<ExcelRowResult<TModel>> ReadExcelStream<TModel>(Stream excelStream, int? headerRow = 1, IEnumerable<IRowValidator<TModel>>? rowValidators = null) where TModel : new()
        {

            var props = Helpers.GetExcelItemProperties<TModel>();
            var mapping = ColumnMapping.BuildColumnMapping<TModel>(props);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(excelStream);
            try
            {
                int currentRowNumber = 0;

                int effectiveHeaderRow = headerRow ?? 1;  
                int firstDataRow = effectiveHeaderRow + 1;

                if (effectiveHeaderRow < 1)
                {
                    throw new ArgumentOutOfRangeException(nameof(headerRow), "Header row must be 1 or greater.");
                }


                do
                {
                    currentRowNumber = 0;

                    while (reader.Read()) 
                    {
                        currentRowNumber++;

                        if (currentRowNumber < firstDataRow) continue;

                        if (Helpers.IsRowEmptyStreaming(reader, mapping.Values.Select(m => m.ColumnIndex)))
                            continue;

                        var model = new TModel();
                        var rowResult = new ExcelRowResult<TModel>(currentRowNumber, model);

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

                            object? raw = Helpers.GetValueSafe(reader, col - 1);

                            try
                            {
                                if (Helpers.IsNullOrEmptyLike(raw))
                                {
                                    item.Value = item.DefaultValue;
                                }
                                else if (item.CustomParse is not null)
                                {
                                    var parseResult = item.CustomParse(raw);

                                    var parsedValue = parseResult.Item1;
                                    var isValid = (bool)parseResult.Item2;
                                    var errMsg = (string?)parseResult.Item3;


                                    if (!isValid)
                                    {
                                        item.Value = item.DefaultValue;
                                        item.SetError(errMsg ?? $"Custom parse failed for {prop.Name}");
                                    }
                                    else
                                    {
                                        item.Value = parsedValue;
                                    }
                                }
                                else
                                {
                                    item.Value = Helpers.ConvertToGeneric(item, raw, prop.PropertyType);
                                }
                            }
                            catch (Exception ex)
                            {
                                item.Value = item.DefaultValue;

                                //TODO : more  friendly error message
                                item.SetError($"Invalid '{prop.Name}' value at row {currentRowNumber}, col {col}: {ex.Message}");
                            }
                        }

                        if (rowValidators != null)
                        {
                            foreach (var validator in rowValidators)
                            {
                                try { validator.Validate(rowResult); }
                                catch (Exception ex) { rowResult.SetError($"Row validator threw: {ex.Message}"); }
                            }
                        }

                        foreach (var prop in mapping)
                        {
                            dynamic? item = prop.Key.GetValue(model);
                            if (item == null) continue;

                            if (!item.IsValid())
                            {
                                rowResult.SetCellError(prop.Key.Name, item.Error!);
                                continue;
                            }

                            item.RunValidators();

                            if (!item.IsValid())
                                rowResult.SetCellError(prop.Key.Name, item.Error!);
                        }

                        yield return rowResult;
                    }

                    break;
                } while (reader.NextResult());
            }
            finally
            {
                reader.Dispose();
            }
        }


    }
}
