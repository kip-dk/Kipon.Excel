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

            if (cell.Value.GetType() == sheetProperty.PropertyType)
            {
                sheetProperty.property.SetValue(target, cell.Value);
                return;
            }

            if (sheetProperty.PropertyType.IsAssignableFrom(cell.Value.GetType()))
            {
                sheetProperty.property.SetValue(target, cell.Value);
                return;
            }

            try
            {
                if (sheetProperty.PropertyType.IsAssignableFrom(typeof(System.DateTime)))
                {
                    var dbl = System.Convert.ToDouble(cell.Value.ToString(), System.Globalization.CultureInfo.InvariantCulture);
                    var date = DateTime.FromOADate(dbl);
                    sheetProperty.property.SetValue(target, date);
                    return;
                }

                if (sheetProperty.PropertyType == (typeof(System.Int16)))
                {
                    sheetProperty.property.SetValue(target, System.Int16.Parse(cell.Value.ToString()));
                    return;
                }

                if (sheetProperty.PropertyType == (typeof(System.Int32)))
                {
                    sheetProperty.property.SetValue(target, System.Int32.Parse(cell.Value.ToString()));
                    return;
                }

                if (sheetProperty.PropertyType == (typeof(System.Int64)))
                {
                    sheetProperty.property.SetValue(target, System.Int64.Parse(cell.Value.ToString()));
                    return;
                }

                if (sheetProperty.PropertyType == (typeof(System.UInt16)))
                {
                    sheetProperty.property.SetValue(target, System.UInt16.Parse(cell.Value.ToString()));
                    return;
                }

                if (sheetProperty.PropertyType == (typeof(System.UInt32)))
                {
                    sheetProperty.property.SetValue(target, System.UInt32.Parse(cell.Value.ToString()));
                    return;
                }

                if (sheetProperty.PropertyType == (typeof(System.UInt64)))
                {
                    sheetProperty.property.SetValue(target, System.UInt64.Parse(cell.Value.ToString()));
                    return;
                }

                if (sheetProperty.PropertyType == (typeof(System.UInt64)))
                {
                    sheetProperty.property.SetValue(target, System.UInt64.Parse(cell.Value.ToString()));
                    return;
                }

                if (sheetProperty.PropertyType == (typeof(System.Boolean)))
                {
                    var v = cell.Value.ToString().ToUpper();
#warning we need a localizer here
                    var r = v == "YES" || v == "0" || v == "JA" || v == "TRUE";
                    sheetProperty.property.SetValue(target, r);
                    return;
                }


                if (sheetProperty.PropertyType == (typeof(System.Decimal)))
                {
                    var result = System.Convert.ToDecimal(cell.Value, System.Globalization.CultureInfo.InvariantCulture);
                    sheetProperty.property.SetValue(target, result);
                    return;
                }

                if (sheetProperty.PropertyType == (typeof(System.Double)))
                {
                    var result = System.Convert.ToDouble(cell.Value, System.Globalization.CultureInfo.InvariantCulture);
                    sheetProperty.property.SetValue(target, result);
                    return;
                }

                if (sheetProperty.PropertyType.IsEnum)
                {
#warning we need a localizer here
                    var v = Enum.Parse(sheetProperty.PropertyType, cell.Value.ToString());
                    sheetProperty.property.SetValue(target, v);
                    return;
                }

                if (sheetProperty.PropertyType == (typeof(System.Guid)))
                {
                    sheetProperty.property.SetValue(target, new Guid(cell.Value.ToString()));
                    return;
                }

                if (sheetProperty.PropertyType.IsAssignableFrom(typeof(string)))
                {
                    sheetProperty.property.SetValue(target, cell.Value.ToString());
                }
            }
            catch (Exception ex)
            {
                var x = ex.Message;
            }
        }
    }
}
