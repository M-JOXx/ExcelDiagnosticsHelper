using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Models
{
    /// <summary>
    /// Holds one imported row (details of row).
    /// </summary>
    public class ExcelRowResult<TModel>
    {
        public int SourceRowNumber { get; }
        public TModel Data { get; }

        // Existing diagnostics
        public Dictionary<string, string> CellErrors { get; } = new(); // PropertyName -> Error
        public List<string> RowErrors { get; } = new();

        public Dictionary<string, List<string>> CellWarnings { get; } = new(); // PropertyName -> warnings
        public List<string> RowWarnings { get; } = new();

        public ExcelRowResult(int sourceRowNumber, TModel data)
        {
            SourceRowNumber = sourceRowNumber;
            Data = data;
        }

        public bool IsValid() => !CellErrors.Any() && !RowErrors.Any();

        public void SetError(string error) => RowErrors.Add(error);
        public void SetCellError(string propertyName, string error)
            => CellErrors[propertyName] = error;


        //warning helpers
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
}
