using Kipon.Excel.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.WriterImplementation.Factories
{
    /// <summary>
    /// From an object instance, resolve the implementation of IEnumable<ISheet> that is relevant
    /// </summary>
    internal class SheetsResolver : TypeCachedResolver<IEnumerable<ISheet>, Models.Sheets.AbstractBaseSheets>
    {
        protected override Models.Sheets.AbstractBaseSheets ResolveType(Type instanceType) 
        {
            #region instance is a single ISheet, simply wrap it in a single sheet
            if (typeof(Kipon.Excel.Api.ISheet).IsAssignableFrom(instanceType))
            {
                return new Kipon.Excel.WriterImplementation.Models.Sheets.SingleSheets();
            }
            #endregion

            // We know for sure that instance IS NOT and IEnumerable<ISheet>, because the AbstractBaseResolver would have returned
            // such instance before hitting this place. So we need to look more into details on what we have

            var properties = instanceType.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

            // Do we have sheet properties, then they are the most important to resolve sheets
            #region sheetsattribute decorated properties
            {
                List<PropertyInfoContainer> decorated = new List<PropertyInfoContainer>();
                int index = 100000;
                foreach (var prop in properties)
                {
                    var attr = (Kipon.Excel.Attributes.SheetAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.SheetAttribute), true).FirstOrDefault();
                    if (attr != null)
                    {
                        var next = new PropertyInfoContainer { PropertyInfo = prop, Sort = attr.Sort ?? index };

                        if (attr.Sort == null)
                        {
                            var sort = (Kipon.Excel.Attributes.SortAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.SortAttribute), true).FirstOrDefault();
                            if (sort != null)
                            {
                                next.Sort = sort.Value;
                            }
                        }
                        decorated.Add(next);
                    }
                }
                if (decorated.Count > 0)
                {
                    var result = new Models.Sheets.PropertySheets();
                    var sheets = (from pc in decorated orderby pc.Sort, pc.PropertyInfo.Name select pc.PropertyInfo).ToArray();
                    result.AddSheetAttributeDecoratedProperties(instanceType, sheets);
                    return result;
                }
            }
            #endregion

            #region ok, no decorations to help us, lets see if we have anything indicating this is a multi sheet thing
            {
                List<PropertyInfoContainer> sheetProperties = new List<PropertyInfoContainer>();
                var sheetResolver = new SheetResolver();
                var index = 100000;
                foreach (var prop in properties)
                {
                    var ignore = (Kipon.Excel.Attributes.IgnoreAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IgnoreAttribute), true).FirstOrDefault();
                    if (ignore != null)
                    {
                        continue;
                    }

                    PropertyInfoContainer next = null;

                    if (typeof(Kipon.Excel.Api.ISheet).IsAssignableFrom(prop.PropertyType))
                    {
                        next = new PropertyInfoContainer { PropertyInfo = prop, Sort = index };
                        sheetProperties.Add(next);
                    }

                    if (next == null && sheetResolver.IsSheet(prop.PropertyType))
                    {
                        next = new PropertyInfoContainer { PropertyInfo = prop, Sort = index };
                        sheetProperties.Add(next);
                    }

                    if (next != null)
                    {
                        var sort = (Kipon.Excel.Attributes.SortAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.SortAttribute), true).FirstOrDefault();
                        if (sort != null)
                        {
                            next.Sort = sort.Value;
                        }
                    }
                }

                if (sheetProperties.Count > 0)
                {
                    var result = new Models.Sheets.PropertySheets();
                    var sheets = (from pc in sheetProperties orderby pc.Sort, pc.PropertyInfo.Name select pc.PropertyInfo).ToArray();
                    result.AddUndecoratedProperties(instanceType, sheets);
                    return result;
                }
            }
            #endregion

            #region single sheet based on array
            {
                if (instanceType.IsArray)
                {
                    return new Kipon.Excel.WriterImplementation.Models.Sheets.SingleSheets();
                }

                if (typeof(System.Collections.IEnumerable).IsAssignableFrom(instanceType) && instanceType.IsGenericType)
                {
                    return new Kipon.Excel.WriterImplementation.Models.Sheets.SingleSheets();
                }
            }
            #endregion

            throw new Kipon.Excel.Exceptions.UnresolveableTypeException(instanceType, typeof(IEnumerable<ISheet>));
        }


        internal class PropertyInfoContainer
        {
            internal int Sort { get; set; }
            internal System.Reflection.PropertyInfo PropertyInfo { get; set; }
        }
    }
}
