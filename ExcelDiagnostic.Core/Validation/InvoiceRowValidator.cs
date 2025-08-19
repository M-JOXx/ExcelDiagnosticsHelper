using ExcelDiagnostic.Core.Contracts;
using ExcelDiagnostic.Core.Models;
using ExcelDiagnostic.Core.Templates;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Validation
{
    public class InvoiceRowValidator : IRowValidator<ExcelInvoiceRow>
    {

        public void Validate(ExcelRowResult<ExcelInvoiceRow> row)
        {
            var data = row.Data;

            // Cells rules
            data.ItemCode.AddValidator(v => !string.IsNullOrWhiteSpace(v) && v!.Length >= 3,
                "ItemCode must have at least 3 characters.");

            data.Amount.AddValidator(v => v.HasValue, "Amount must be provided.");

            data.OrderId.AddValidator(v => v >= 5, "OrderId must be >= 5");


            // Row level business rule
            var type = (data.Type.Value ?? string.Empty).Trim();
            var amount = data.Amount.Value;

            if (string.Equals(type, "Refund", StringComparison.OrdinalIgnoreCase))
            {
                if (amount <= 0)
                    row.SetError("For 'Refund' rows, Amount must be positive.");
            }

            //  Warning 
            if (data.Amount.Value > 500)
            {
                row.SetCellWarning(nameof(data.Amount), "Amount is unusually high.");
            }
        }
    }
}
