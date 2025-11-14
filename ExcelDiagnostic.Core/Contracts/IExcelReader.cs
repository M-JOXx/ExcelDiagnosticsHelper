using ExcelDiagnostic.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Contracts
{
    /// <summary>
    /// Interface for Excel readers
    /// </summary>
    public interface IExcelReader
    {
        /// <summary>
        /// Reads an Excel file and returns diagnostic results.
        /// TModel must implement IExcelSheet interface.
        /// </summary>
        ExcelDiagnosticResult<TModel> ReadExcel<TModel>(
            string filePath,
            int? headerRow = 1,
            IEnumerable<IRowValidator<TModel>>? rowValidators = null)
            where TModel : IExcelSheet, new();

        /// <summary>
        /// Reads an Excel stream and returns diagnostic results.
        /// TModel must implement IExcelSheet interface.
        /// </summary>
        ExcelDiagnosticResult<TModel> ReadExcel<TModel>(
            Stream excelStream,
            int? headerRow = 1,
            IEnumerable<IRowValidator<TModel>>? rowValidators = null)
            where TModel : IExcelSheet, new();
    }
}