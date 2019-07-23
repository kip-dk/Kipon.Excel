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

            var cellsResolver = new Kipon.Excel.Implementation.Factories.CellsResolver();

            foreach (var prop in publicProperties)
            {
                if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                {
                    var innerType = prop.PropertyType.GetGenericArguments()[0];
                    if (cellsResolver.HasCells(innerType))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
