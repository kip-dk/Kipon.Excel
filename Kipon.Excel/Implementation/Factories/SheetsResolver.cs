using Kipon.Excel.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
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
                return new Kipon.Excel.Implementation.Models.Sheets.SingleSheets();
            }
            #endregion

            // We know for sure that instance IS NOT and IEnumerable<ISheet>, because the AbstractBaseResolver would have returned
            // such instance before hitting this place. So we need to look more into details on what we have

            var properties = instanceType.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

            // Do we have sheet properties, then they are the most important to resolve sheets
            #region sheetsattribute decorated properties
            {
                List<System.Reflection.PropertyInfo> decorated = new List<System.Reflection.PropertyInfo>();
                foreach (var prop in properties)
                {
                    var attr = prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.SheetAttribute), true);
                    if (attr != null && attr.Length > 0)
                    {
                        decorated.Add(prop);
                    }
                }
                if (decorated.Count > 0)
                {
                    var result = new Models.Sheets.PropertySheets();
                    result.AddSheetAttributeDecoratedProperties(instanceType, decorated.ToArray());
                    return result;
                }
            }
            #endregion

            #region ok, no decorations to help us, lets see if we have anything indicating this is a multi sheet thing
            {
                List<System.Reflection.PropertyInfo> sheetProperties = new List<System.Reflection.PropertyInfo>();
                var sheetResolver = new SheetResolver();
                foreach (var prop in properties)
                {
                    if (typeof(Kipon.Excel.Api.ISheet).IsAssignableFrom(prop.PropertyType))
                    {
                        sheetProperties.Add(prop);
                        continue;
                    }

                    if (sheetResolver.IsSheet(prop.PropertyType))
                    {
                        sheetProperties.Add(prop);
                        continue;
                    }
                }
                if (sheetProperties.Count > 0)
                {
                    var result = new Models.Sheets.PropertySheets();
                    result.AddUndecoratedProperties(instanceType, sheetProperties.ToArray());
                    return result;
                }
            }
            #endregion

            #region single sheet based on array
            {
                if (instanceType.IsArray)
                {
                    return new Kipon.Excel.Implementation.Models.Sheets.SingleSheets();
                }

                if (typeof(System.Collections.IEnumerable).IsAssignableFrom(instanceType) && instanceType.IsGenericType)
                {
                    return new Kipon.Excel.Implementation.Models.Sheets.SingleSheets();
                }
            }
            #endregion

            throw new Kipon.Excel.Exceptions.UnresolveableTypeException(instanceType, typeof(IEnumerable<ISheet>));
        }
    }
}
