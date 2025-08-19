using ExcelDiagnostic.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Mapping
{
  
    internal static class ColumnMapping
    {
        /// <summary>
        /// Builds a mapping of PropertyInfo to ColumnMeta for a given model type. give more details of props 
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="props"></param>
        /// <returns></returns>
        public static Dictionary<PropertyInfo, ColumnMeta> BuildColumnMapping<TModel>(IEnumerable<PropertyInfo> props)
        {
            var mapping = new Dictionary<PropertyInfo, ColumnMeta>();
            int running = 1;
            foreach (var p in props)
            {
                dynamic? item = Activator.CreateInstance(p.PropertyType);
                int colIndex = item?.Index ?? running++;
                mapping[p] = new ColumnMeta { ColumnIndex = colIndex };
            }
            return mapping;
        }


        public sealed class ColumnMeta
        {
            public int ColumnIndex { get; init; }
        }

    }

}
