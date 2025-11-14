using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Models
{
    /// <summary>
    /// Wrapper for a cell. Holds column index, parsed value, default value,
    /// parse function, validators, and error message. Provides IsValid/SetError.
    /// </summary>
    public class ExcelCell<T>
    {
        // it could be set by order of model(template)
        /// <summary> column index in the worksheet.
        /// </summary>
        public int? Index { get; set; }

        /// <summary> Parsed value. </summary>
        public T? Value { get; set; }

        /// <summary> Default value used when the cell is empty. </summary>
        public T? DefaultValue { get; set; }

        //Custom parse function that takes raw object and returns a tuple of parsed value, validity, and error message ,
        //i make it more detailed than you mention in task and  one example in validator (amount) .
        public Func<object?, (T? ParsedValue, bool IsValid, string? ErrorMessage)>? CustomParse { get; set; }

        /// <summary> Error message for this cell (null/empty if valid). </summary>
        public string? Error { get; private set; }

        /// <summary> List of warnings for this cell. </summary>
        public List<string> Warnings { get; } = new();

        // Reference to parent row result for auto-sync
        internal object? ParentRow { get; set; }
        internal string? PropertyName { get; set; }

        public bool IsValid() => string.IsNullOrWhiteSpace(Error);

        /// <summary>
        /// Sets an error on this cell. Automatically syncs to parent row.
        /// Example: row.Data.Amount.SetCellError("Amount must be positive");
        /// </summary>
        public void SetCellError(string error)
        {
            Error = error;
            // Auto-sync to parent row if available
            if (ParentRow != null && PropertyName != null)
            {
                ParentRow.GetType()
                    .GetMethod("SetCellError", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    ?.Invoke(ParentRow, new object[] { PropertyName, error });
            }
        }

        /// <summary>
        /// Adds a warning to this cell. Automatically syncs to parent row.
        /// Example: row.Data.Amount.SetCellWarning("Amount is unusually high");
        /// </summary>
        public void SetCellWarning(string warning)
        {
            Warnings.Add(warning);
            // Auto-sync to parent row if available
            if (ParentRow != null && PropertyName != null)
            {
                ParentRow.GetType()
                    .GetMethod("SetCellWarning", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    ?.Invoke(ParentRow, new object[] { PropertyName, warning });
            }
        }

        // Internal method for parsing - keep for backward compatibility
        internal void SetError(string error) => Error = error;

        //--------

        /// <summary> Optional custom validators that must all pass. </summary>
        public List<Func<T?, bool>> Validators { get; } = new();

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
}