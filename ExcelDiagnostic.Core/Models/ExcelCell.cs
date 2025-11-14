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
        public string? Error { get; private set; }
        public bool IsValid() => string.IsNullOrWhiteSpace(Error);

        public void SetError(string error) => Error = error;

        //--------
        
        /// <summary> Optional custom validators that must all pass. </summary>
        public List<Func<T?, bool>> Validators { get; } = new();

        /// <summary> Error message for this cell (null/empty if valid). </summary>

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
