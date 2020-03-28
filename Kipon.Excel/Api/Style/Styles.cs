using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace Kipon.Excel.Api.Style
{
    public static class Styles
    {
        public static readonly UInt32Value ODD_DATE_LOCKED = (UInt32Value)9;
        public static readonly UInt32Value EVEN_DATE_LOCKED = (UInt32Value)8;

        public static readonly UInt32Value ODD_DATE_UNLOCKED = (UInt32Value)7;
        public static readonly UInt32Value EVEN_DATE_UNLOCKED = (UInt32Value)6;

        public static readonly UInt32Value BOLD_STYLE_INDEX = (UInt32Value)5;

        public static readonly UInt32Value ODD_STYLE_INDEX_LOCKED = (UInt32Value)4;
        public static readonly UInt32Value EVEN_STYLE_INDEX_LOCKED = (UInt32Value)3;

        public static readonly UInt32Value ODD_STYLE_INDEX_UNLOCKED = (UInt32Value)2;
        public static readonly UInt32Value EVEN_STYLE_INDEX_UNLOCKED = (UInt32Value)1;

        public static System.UInt32 Resolve(Kipon.Excel.Api.ICell cell, bool isEven)
        {
            var row = cell.Coordinate.Point.Last();

            if (row == 0) return Kipon.Excel.Api.Style.Styles.BOLD_STYLE_INDEX;

            var isreadonly = false;
            var readonlyimpl = cell as Kipon.Excel.Api.Cell.IReadonly;
            if (readonlyimpl != null)
            {
                isreadonly = readonlyimpl.IsReadonly;
            }

            var datatype = cell as Kipon.Excel.Api.Cell.IDataType;
            if (datatype != null && datatype.ValueType != null && datatype.ValueType == typeof(DateTime))
            {
                if (isEven && !isreadonly)
                {
                    return Kipon.Excel.Api.Style.Styles.EVEN_DATE_UNLOCKED;
                }

                if (isEven && isreadonly)
                {
                    return Kipon.Excel.Api.Style.Styles.EVEN_DATE_LOCKED;
                }

                if (!isreadonly)
                {
                    return Kipon.Excel.Api.Style.Styles.ODD_DATE_UNLOCKED;
                }
                return Kipon.Excel.Api.Style.Styles.ODD_DATE_LOCKED;
            }

            if (isEven && !isreadonly)
            {
                return Kipon.Excel.Api.Style.Styles.EVEN_STYLE_INDEX_UNLOCKED;
            }

            if (isEven && isreadonly)
            {
                return Kipon.Excel.Api.Style.Styles.EVEN_STYLE_INDEX_LOCKED;
            }

            if (!isreadonly)
            {
                return Kipon.Excel.Api.Style.Styles.ODD_STYLE_INDEX_UNLOCKED;
            }

            return Kipon.Excel.Api.Style.Styles.ODD_STYLE_INDEX_LOCKED;
        }
    }
}
