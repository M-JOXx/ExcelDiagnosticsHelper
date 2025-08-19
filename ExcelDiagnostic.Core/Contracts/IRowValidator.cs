using ExcelDiagnostic.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Contracts
{
    /// <summary>
    /// Optional contract for row-level validators or cell errors .
    /// </summary>
    public interface IRowValidator<TModel>
    {
        void Validate(ExcelRowResult<TModel> row);
    }

}
