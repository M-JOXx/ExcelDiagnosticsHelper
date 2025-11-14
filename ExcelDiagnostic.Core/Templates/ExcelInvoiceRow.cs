using ExcelDiagnostic.Core.Contracts;
using ExcelDiagnostic.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Templates
{
    //which is the model i call it templates so you can add more than one hre.
    public class ExcelInvoiceRow : IExcelSheet
    {
        public ExcelCell<int> OrderId { get; set; } = new()
        {
            Index = 1,
            Required = true
        };

        public ExcelCell<string> ItemCode { get; set; } = new()
        {
            Index = 2,
            Required = true
        };

        public ExcelCell<decimal?> Amount { get; set; } = new()
        {
            Index = 3,
            CustomParse = raw =>
            {
                if (raw == null) return (null, false, "Amount empty");
                var s = raw.ToString()?.Trim();
                // accept comma or dot as decimal separator
                s = s?.Replace(',', '.');
                if (!decimal.TryParse(s, out var d))
                    return (null, false, $"Amount '{s}' is not a valid decimal");
                return (d, true, null);
            },
        };

        public ExcelCell<string> Type { get; set; } = new()
        {
            Index = 4,
            Required = true
        };

        public ExcelCell<string> CustomerName { get; set; } = new()
        {
            Index = 5,
            Required = true
        };

    }

}
