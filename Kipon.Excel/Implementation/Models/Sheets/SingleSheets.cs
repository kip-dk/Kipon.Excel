using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Models.Sheets
{
    internal class SingleSheets : AbstractBaseSheets
    {
        public override void Populate(object instance)
        {
            var sheetResolver = new Kipon.Excel.Implementation.Factories.SheetResolver();

            var sheet = sheetResolver.Resolve(instance);
            this.Add(sheet);

            var propertySheet = sheet as Kipon.Excel.Implementation.Models.Sheet.AbstractBaseSheet;
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
