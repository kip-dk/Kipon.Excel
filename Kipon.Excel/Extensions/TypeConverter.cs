using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace Kipon.Excel.Extensions
{
    internal static class TypeConverter
    {
        public static object ToValue(this object value, Type propertyType)
        {
            if (value == null)
            {
                return null;
            }

            var nullType = Nullable.GetUnderlyingType(propertyType);
            if (nullType != null)
            {
                propertyType = nullType;
            }

            var valueType = value.GetType();

            if (valueType == propertyType)
            {
                return value;
            }

            if (propertyType.IsAssignableFrom(valueType))
            {
                return value;
            }

            if (propertyType.IsAssignableFrom(typeof(System.DateTime)))
            {
                var dbl = System.Convert.ToDouble(value.ToString(), System.Globalization.CultureInfo.InvariantCulture);
                var date = DateTime.FromOADate(dbl);
                return date;
            }

            if (propertyType == (typeof(System.Int16)))
            {
                return Int16.Parse(value.ToString());
            }

            if (propertyType == (typeof(System.Int32)))
            {
                return Int32.Parse(value.ToString());
            }

            if (propertyType == (typeof(System.Int64)))
            {
                return Int64.Parse(value.ToString());
            }

            if (propertyType == (typeof(System.UInt16)))
            {
                return System.UInt16.Parse(value.ToString());
            }

            if (propertyType == (typeof(System.UInt32)))
            {
                return System.UInt32.Parse(value.ToString());
            }

            if (propertyType == (typeof(System.UInt64)))
            {
                return System.UInt64.Parse(value.ToString());
            }

            if (propertyType == (typeof(System.Boolean)))
            {
                var v = value.ToString().ToUpper();
#warning we need a localizer here
                var r = v == "YES" || v == "0" || v == "JA" || v == "TRUE";
                return r;
            }


            if (propertyType == (typeof(System.Decimal)))
            {
                var result = System.Convert.ToDecimal(value, System.Globalization.CultureInfo.InvariantCulture);
                return result;
            }

            if (propertyType == (typeof(System.Double)))
            {
                var result = System.Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
                return result;
            }

            if (propertyType.IsEnum)
            {
#warning we need a localizer here
                var v = Enum.Parse(propertyType, value.ToString());
                return v;
            }

            if (propertyType == (typeof(System.Guid)))
            {
                return new Guid(value.ToString());
            }

            if (propertyType.IsAssignableFrom(typeof(string)))
            {
                return value.ToString();
            }
            return null;
        }
    }
}
