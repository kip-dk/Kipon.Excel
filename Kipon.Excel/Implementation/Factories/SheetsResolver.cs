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
    /// <typeparam name="J">J is the actual implementation of IEnumerable<ISheet>. It must have a constructor that take 0 arguments</typeparam>
    internal class SheetsResolver<J> : TypeCachedResolver<IEnumerable<ISheet>, Models.Sheets.AbstractBaseSheets>
        where J : IPopulator
    {
        protected override Models.Sheets.AbstractBaseSheets ResolveType(Type instanceType) 
        {
            // We know for sure that instance IS NOT and IEnumerable<ISheet>, because the AbstractBaseResolver would have returned
            // such instance before hitting this place. So we need to look more into details on what we have

            var properties = instanceType.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

            // Do we have sheet properties, then they are the most important to resolve sheets
            #region sheetsattribute decorated
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

            return null;
        }
    }
}
