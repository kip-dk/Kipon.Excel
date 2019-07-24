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
            if (instanceType.IsArray)
            {
                var result = new Kipon.Excel.Implementation.Models.Sheet.PropertySheet();
#warning find relevant properties and parse it to the PropertySheet instance
                return result;
            }

            if (instanceType.IsGenericType && typeof(System.Collections.IEnumerable).IsAssignableFrom(instanceType))
            {
                var result = new Kipon.Excel.Implementation.Models.Sheet.PropertySheet();
#warning find relevant properties and parse it to the PropertySheet instance
                return result;
            }
            return null;
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

            var cellsResolver = new Kipon.Excel.Implementation.Factories.CellsResolver();
            if (type.IsArray)
            {
                var innerType = type.GetElementType();
                return cellsResolver.HasCells(innerType);
            }

            if (type.IsGenericType && typeof(System.Collections.IEnumerable).IsAssignableFrom(type))
            {
                var innerType = type.GetGenericArguments()[0];
                return cellsResolver.HasCells(innerType);
            }

            return false;
        }
    }
}
