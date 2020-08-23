using DocumentFormat.OpenXml.Spreadsheet;
using Kipon.Excel.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.ReaderImplementation.Converters
{
    internal class PropertyCellConverter
    {
        private object target;
        internal PropertyCellConverter(object target)
        {
            this.target = target;
        }

        internal void Convert(Kipon.Excel.Reflection.PropertySheet.SheetProperty sheetProperty, Kipon.Excel.Api.ICell cell)
        {
            if (cell.Value == null)
            {
                return;
            }

            if (!sheetProperty.canWrite)
            {
                return;
            }

            if (sheetProperty.indexExpression != null)
            {
                // this is an indexer
                if (sheetProperty.property.GetValue(target) == null)
                {
                    var dic = Activator.CreateInstance(sheetProperty.PropertyType);
                    sheetProperty.property.SetValue(target, dic);
                }
                var index = sheetProperty.property.GetValue(target);
                var name = sheetProperty[cell.Coordinate.Point.First()];
                try
                {
                    var value = cell.Value.ToValue(sheetProperty.indexType);

                    if (value != null)
                    {
                        sheetProperty.indexAddMethod.Invoke(index, new object[] { name, value });
                    }
                } catch (Exception)
                {
                    // catch by design to allow an excel to parse even if it has data in celle that are not compatible, they will just be null
                }
                return;
            }

            {
                object value = null;
                try
                {
                    value = sheetProperty.ValueOf(cell.Value);

                    if (value != null)
                    {
                        sheetProperty.property.SetValue(target, value);
                        return;
                    }
                }
                catch (Exception)
                {
                    // catch by design to allow an excel to parse even if it has data in celle that are not compatible, they will just be null
                }
            }
        }
    }
}
