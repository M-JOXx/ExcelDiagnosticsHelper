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

        // Internal diagnostics storage
        internal Dictionary<string, string> CellErrors { get; } = new();
        internal Dictionary<string, List<string>> CellWarnings { get; } = new();

        // Public diagnostics - row-level only
        public List<string> RowErrors { get; } = new();
        public List<string> RowWarnings { get; } = new();

        public ExcelRowResult(int sourceRowNumber, TModel data)
        {
            SourceRowNumber = sourceRowNumber;
            Data = data;

            // Wire up parent references for auto-sync
            LinkCellsToRow();
        }

        /// <summary>
        /// Links all ExcelCell properties to this row for automatic syncing
        /// </summary>
        private void LinkCellsToRow()
        {
            if (Data == null) return;

            var props = Data.GetType().GetProperties();
            foreach (var prop in props)
            {
                if (prop.PropertyType.IsGenericType &&
                    prop.PropertyType.GetGenericTypeDefinition() == typeof(ExcelCell<>))
                {
                    var cell = prop.GetValue(Data);
                    if (cell != null)
                    {
                        // Use reflection to set parent reference and property name
                        var parentRowProp = prop.PropertyType.GetProperty("ParentRow",
                            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public);
                        var propertyNameProp = prop.PropertyType.GetProperty("PropertyName",
                            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public);

                        parentRowProp?.SetValue(cell, this);
                        propertyNameProp?.SetValue(cell, prop.Name);
                    }
                }
            }
        }

        public bool IsValid() => !CellErrors.Any() && !RowErrors.Any();

        /// <summary>
        /// Sets a row-level error (not specific to any column).
        /// Example: row.SetError("OrderId and Amount must both be above 5");
        /// </summary>
        public void SetError(string error) => RowErrors.Add(error);

        /// <summary>
        /// Sets a row-level warning (not specific to any column).
        /// Example: row.SetRowWarning("This row needs manual review");
        /// </summary>
        public void SetRowWarning(string warning) => RowWarnings.Add(warning);

        // Internal methods - used by reader during parsing and by ExcelCell for auto-sync
        internal void SetCellError(string propertyName, string error)
            => CellErrors[propertyName] = error;

        internal void SetCellWarning(string propertyName, string warning)
        {
            if (!CellWarnings.TryGetValue(propertyName, out var list))
            {
                list = new List<string>();
                CellWarnings[propertyName] = list;
            }
            list.Add(warning);
        }

        /// <summary>
        /// Syncs errors and warnings from ExcelCell properties to internal dictionaries.
        /// This is called automatically by the reader after validators run.
        /// You normally don't need to call this manually since cells auto-sync now.
        /// </summary>
        public void SyncCellDiagnostics()
        {
            if (Data == null) return;

            var props = Data.GetType().GetProperties();
            foreach (var prop in props)
            {
                // Check if property is an ExcelCell<T>
                if (prop.PropertyType.IsGenericType &&
                    prop.PropertyType.GetGenericTypeDefinition() == typeof(ExcelCell<>))
                {
                    var cell = prop.GetValue(Data);
                    if (cell == null) continue;

                    // Use reflection to access Error and Warnings
                    var errorProp = prop.PropertyType.GetProperty("Error");
                    var warningsProp = prop.PropertyType.GetProperty("Warnings");

                    var error = errorProp?.GetValue(cell) as string;
                    var warnings = warningsProp?.GetValue(cell) as List<string>;

                    // Sync error
                    if (!string.IsNullOrWhiteSpace(error))
                    {
                        CellErrors[prop.Name] = error;
                    }

                    // Sync warnings
                    if (warnings != null && warnings.Count > 0)
                    {
                        if (!CellWarnings.TryGetValue(prop.Name, out var list))
                        {
                            list = new List<string>();
                            CellWarnings[prop.Name] = list;
                        }

                        // Add only new warnings (avoid duplicates from auto-sync)
                        foreach (var warning in warnings)
                        {
                            if (!list.Contains(warning))
                            {
                                list.Add(warning);
                            }
                        }
                    }
                }
            }
        }
    }
}