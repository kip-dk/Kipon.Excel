using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Reflection
{
    internal class PropertyCell
    {
        private static readonly Type[] SUPPORTED_SHEET_PROPERTY_TYPES = new Type[]
        {
            typeof(System.Int16),
            typeof(System.Int32),
            typeof(System.Int64),
            typeof(System.UInt16),
            typeof(System.UInt32),
            typeof(System.UInt64),
            typeof(System.Boolean),
            typeof(System.Decimal),
            typeof(System.Double),
            typeof(System.Single),
            typeof(System.Enum),
            typeof(System.String),
            typeof(System.DateTime),
            typeof(System.Guid)
        };


        internal static bool HasCell(Type type)
        {
            var properties = type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

            foreach (var prop in properties)
            {
                var columnAttr = prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.ColumnAttribute), true).FirstOrDefault();
                if (columnAttr != null)
                {
                    return true;
                }

                var indexProperty = type.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IndexColumnAttribute), false).FirstOrDefault();
                if (indexProperty != null)
                {
                    return true;
                }

                if (PropertyCell.IsCell(prop.PropertyType))
                {
                    return true;
                }
            }
            return false;
        }

        internal static bool IsCell(Type type)
        {
            if (SUPPORTED_SHEET_PROPERTY_TYPES.Contains(type))
            {
                return true;
            }

            if (type.IsEnum)
            {
                return true;
            }

            var nullableType = Nullable.GetUnderlyingType(type);
            if (nullableType != null && SUPPORTED_SHEET_PROPERTY_TYPES.Contains(nullableType))
            {
                return true;
            }

            if (nullableType != null && nullableType.IsEnum)
            {
                return true;
            }

            return false;
        }
    }
}
