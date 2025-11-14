using ExcelDiagnostic.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Contracts
{
    public interface IExcelReader
    {
        ExcelDiagnosticResult<TModel> ReadExcel<TModel>(
            string filePath,
            int? headerRow = 1,
            IEnumerable<IRowValidator<TModel>>? rowValidators = null) where TModel : new();

        ExcelDiagnosticResult<TModel> ReadExcel<TModel>(
            Stream excelStream,
            int? headerRow = 1,
            IEnumerable<IRowValidator<TModel>>? rowValidators = null) where TModel : new();


    }
}
