using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Implementation.Models.Sheet;

namespace Kipon.Excel.Implementation.Factories
{
    internal class SheetResolver : TypeCachedResolver<Kipon.Excel.Api.ISheet, Models.Sheet.AbstractBaseSheet>
    {
        private static readonly Type[] SUPPORTED_SHEET_PROPERTY_TYPES = new Type[]
        {
            typeof(System.Int16),
            typeof(System.Int32),
            typeof(System.Int64),
            typeof(System.IntPtr),
            typeof(System.UInt16),
            typeof(System.UInt32),
            typeof(System.UInt64),
            typeof(System.UIntPtr),
            typeof(System.Boolean),
            typeof(System.Decimal),
            typeof(System.Double),
            typeof(System.Enum),
            typeof(System.String),
            typeof(System.DateTime)
        };

        protected override AbstractBaseSheet ResolveType(Type instanceType)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Will return true if the type can be resolved to a Sheet through public properties, otherwise false
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        internal bool IsSheet(Type type)
        {
            if (type.IsPrimitive)
            {
                return false;
            }

            if (type.IsEnum)
            {
                return false;
            }

            if (type == typeof(string))
            {
                return false;
            }


            var publicProperties = type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            if (publicProperties == null || publicProperties.Length == 0)
            {
                return false;
            }

            foreach (var prop in publicProperties)
            {
                if (SUPPORTED_SHEET_PROPERTY_TYPES.Contains(prop.PropertyType))
                {
                    return true;
                }

                if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    var innerType = prop.PropertyType.GetGenericArguments()[0];
                    if (SUPPORTED_SHEET_PROPERTY_TYPES.Contains(innerType))
                    {
                        return true;
                    }
                }
            }
            return true;
        }
    }
}
