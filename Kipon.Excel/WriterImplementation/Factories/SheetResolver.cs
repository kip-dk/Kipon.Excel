using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.WriterImplementation.Models.Sheet;

namespace Kipon.Excel.WriterImplementation.Factories
{
    internal class SheetResolver : TypeCachedResolver<Kipon.Excel.Api.ISheet, Models.Sheet.AbstractBaseSheet>
    {
        protected override AbstractBaseSheet ResolveType(Type instanceType)
        {
            if (instanceType.IsArray)
            {
                var result = new Kipon.Excel.WriterImplementation.Models.Sheet.PropertySheet();
                var elementType = instanceType.GetElementType();
                var elementProperties = this.FindProperties(elementType);
                result.AddPropertyInfos(elementType, elementProperties);
                return result;
            }

            if (instanceType.IsGenericType && typeof(System.Collections.IEnumerable).IsAssignableFrom(instanceType))
            {
                var result = new Kipon.Excel.WriterImplementation.Models.Sheet.PropertySheet();
                var elementType = instanceType.GetGenericArguments()[0];
                var elementProperties = this.FindProperties(elementType);
                result.AddPropertyInfos(elementType, elementProperties);
                return result;
            }

            throw new Kipon.Excel.Exceptions.UnresolveableTypeException(instanceType, typeof(Kipon.Excel.Api.ISheet));
        }

        private System.Reflection.PropertyInfo[] FindProperties(Type t)
        {
            var properties = t.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            List<System.Reflection.PropertyInfo> result = new List<System.Reflection.PropertyInfo>();

            #region column decorated properties
            {
                foreach (var prop in properties)
                {
                    var columnAttr = prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.ColumnAttribute), true);
                    if (columnAttr != null && columnAttr.Length > 0)
                    {

                        var ignore = prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IgnoreAttribute), false).Any();
                        if (ignore)
                        {
                            throw new Exceptions.InconsistantPropertyException(prop);
                        }

                        if (prop.GetGetMethod() == null)
                        {
                            throw new Exceptions.InconsistantPropertyException(prop);
                        }

                        result.Add(prop);
                    }
                }
                if (result.Count > 0)
                {
                    return result.ToArray();
                }
            }
            #endregion

            #region duck type properties
            {
                var cellsResolver = new Kipon.Excel.WriterImplementation.Factories.CellsResolver();
                foreach (var prop in properties)
                {
                    var ignore = prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IgnoreAttribute), false).Any();
                    if (ignore)
                    {
                        continue;
                    }

                    if (prop.GetGetMethod() == null)
                    {
                        continue;
                    }

                    if (cellsResolver.IsCell(prop.PropertyType))
                    {
                        result.Add(prop);
                    }
                }
                if (result.Count > 0)
                {
                    return result.ToArray();
                }
            }
            #endregion

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

            var cellsResolver = new Kipon.Excel.WriterImplementation.Factories.CellsResolver();
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
