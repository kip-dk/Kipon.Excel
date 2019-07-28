using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using Kipon.Excel.Api;

namespace Kipon.Excel.Implementation.OpenXml
{
    internal class DataTypeResolver : Kipon.Excel.Implementation.Serialization.IDataTypeResolver
    {
        public int? NumberOfDecimals(ICell cell)
        {
            var decI = cell as Kipon.Excel.Api.Cell.IDecimals;
            if (decI != null)
            {
                return decI.Decimals;
            }
            return null;
        }

        public CellValues Resolve(ICell cell)
        {
            var datatype = cell as Kipon.Excel.Api.Cell.IDataType;
            Type t = typeof(string);

            if (datatype != null && datatype.ValueType != null)
            {
                t = datatype.ValueType;
            }

            if ((datatype == null || datatype.ValueType == null) && cell.Value != null)
            {
                t = cell.Value.GetType();
            }

            var nullableType = Nullable.GetUnderlyingType(t);

            if (nullableType != null)
            {
                t = nullableType;
            }

            if (t == typeof(short) || t == typeof(int) || t == typeof(long) || t == typeof(decimal) || t == typeof(double) || t == typeof(float)  || t == typeof(uint))
            {
                return CellValues.Number;
            }

            if (t == typeof(DateTime))
            {
                return CellValues.Date;
            }

            if (t == typeof(bool))
            {
                return CellValues.Boolean;
            }

            return CellValues.String;
        }
    }
}
