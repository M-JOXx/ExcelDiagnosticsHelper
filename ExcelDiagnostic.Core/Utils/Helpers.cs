using ExcelDiagnostic.Core.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Utils
{
    public class Helpers
    {

        public static bool IsRowEmptyStreaming(IDataReader reader, IEnumerable<int> columns1Based)
        {
            foreach (var c1 in columns1Based)
            {
                int c0 = c1 - 1; 
                if (c0 >= 0 && c0 < reader.FieldCount)
                {
                    var val = reader.GetValue(c0);
                    if (!IsNullOrEmptyLike(val)) return false;
                }
            }
            return true;
        }

        public static object? GetValueSafe(IDataReader reader, int zeroBasedCol)
        {
            if (zeroBasedCol < 0 || zeroBasedCol >= reader.FieldCount) return null;
            var v = reader.GetValue(zeroBasedCol);
            return v == DBNull.Value ? null : v;
        }

        public static bool IsNullOrEmptyLike(object? value)
        {
            if (value is null || value == DBNull.Value) return true;
            if (value is string s) return string.IsNullOrWhiteSpace(s);
            return false;
        }


        public static object? ConvertToGeneric(dynamic excelItem, object? raw, Type excelItemType)
        {
            Type t = excelItemType.GetGenericArguments()[0];
            if (raw is null) return null;

            if (t == typeof(string))
                return raw.ToString()?.Trim();

            var underlying = Nullable.GetUnderlyingType(t);
            if (underlying != null)
            {
                if (IsNullOrEmptyLike(raw)) return null;
                return ChangeType(raw, underlying);
            }

            return ChangeType(raw, t);
        }

        public static object? ChangeType(object value, Type targetType)
        {
            try
            {
                if (targetType == typeof(Guid))
                {
                    if (value is Guid g) return g;
                    return Guid.Parse(value.ToString()!);
                }

                if (targetType.IsEnum)
                {
                    if (value is string es) return Enum.Parse(targetType, es, ignoreCase: true);
                    return Enum.ToObject(targetType, Convert.ChangeType(value, Enum.GetUnderlyingType(targetType)));
                }

                var converter = TypeDescriptor.GetConverter(targetType);
                if (converter.CanConvertFrom(value.GetType()))
                    return converter.ConvertFrom(value);

                return Convert.ChangeType(value, targetType);
            }
            catch
            {
                throw new InvalidCastException($"Cannot convert '{value}' to {targetType.Name}");
            }
        }


        public static IEnumerable<PropertyInfo> GetExcelItemProperties<TModel>()
            => typeof(TModel)
               .GetProperties(BindingFlags.Public | BindingFlags.Instance)
               .Where(p => p.PropertyType.IsGenericType &&
                           p.PropertyType.GetGenericTypeDefinition() == typeof(ExcelCell<>))
               .OrderBy(OrderByDeclarationOrder);

        public static int OrderByDeclarationOrder(PropertyInfo p) => p.MetadataToken;



    }

}
