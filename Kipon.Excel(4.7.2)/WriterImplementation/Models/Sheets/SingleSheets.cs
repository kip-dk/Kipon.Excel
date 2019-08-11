using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.WriterImplementation.Models.Sheets
{

    /// <summary>
    /// An implementation i IEnuabrable&lt;ISeets&gt; based on a single list or array, mapping to a single sheet.
    /// </summary>
    internal class SingleSheets : AbstractBaseSheets
    {
        public override void Populate(object instance)
        {
            if (instance is Kipon.Excel.Api.ISheet)
            {
                this.Add((Kipon.Excel.Api.ISheet)instance);
                return;
            }

            var sheetResolver = new Kipon.Excel.WriterImplementation.Factories.SheetResolver();

            var sheet = sheetResolver.Resolve(instance);
            this.Add(sheet);

            var propertySheet = sheet as Kipon.Excel.WriterImplementation.Models.Sheet.AbstractBaseSheet;
            if (propertySheet != null)
            {
                var type = instance.GetType();
                if (type.IsArray)
                {
                    propertySheet.Title = type.GetElementType().Name;
                    return;
                }

                if (typeof(System.Collections.IEnumerable).IsAssignableFrom(type) && type.IsGenericType)
                {
                    propertySheet.Title = type.GetGenericArguments()[0].Name;
                    return;
                }

                propertySheet.Title = "Sheet 1";
            }
        }
    }
}
