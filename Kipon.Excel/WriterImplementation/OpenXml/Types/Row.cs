using System;
using System.Runtime.Serialization;
using System.Security.Permissions;

namespace Kipon.Excel.WriterImplementation.OpenXml.Types
{
    /// <summary>
    /// A row is basically a simple integer, but limited to the min max rows supported for excel.
    /// Row index start at 0, and limits to 1048576 (-1)
    /// Row has been autoboxed with int, so any valid int value can be casted to a Row, vise versa
    /// and any interface taking a Row as parameter, will also allow an int.
    /// 
    /// the purpose of the type is to inforce correct line numbers
    /// </summary>
    internal struct Row 
    {
        #region fields and consts
        public const int EXCEL_TOTAL_MAXROWS = 1048576;

        private int _value;
        #endregion

        #region statics
        private static readonly System.Collections.Generic.Dictionary<int, Row> _rows = new System.Collections.Generic.Dictionary<int, Row>();

        internal static Row getRow(int value)
        {
            if (_rows.ContainsKey(value))
            {
                return _rows[value];
            }
            var next = new Row(value);
            _rows[value] = next;
            return next;
        }
        #endregion

        #region constructor
        private Row(int value)
        {
            if (value < 0) throw new ArgumentOutOfRangeException("value cannot be less than zero");
            if (value >= EXCEL_TOTAL_MAXROWS) throw new ArgumentOutOfRangeException($"value cannot be greater or equal than {EXCEL_TOTAL_MAXROWS}");
            this._value = value;
        }
        #endregion

        #region properties
        public int Value
        {
            get
            {
                return _value;
            }
        }
        #endregion


        #region overrides
        public override bool Equals(object obj)
        {
            if (obj is Row)
            {
                return this.Value == ((Row)obj).Value;
            }

            if (obj is int)
            {
                return this.Value == (int)obj;
            }

            if (obj is int? && obj != null)
            {
                return this.Value == ((int?)obj).Value;
            }

            return false;
        }
        public override int GetHashCode()
        {
            return this.Value.GetHashCode();
        }

        public override string ToString()
        {
            return this.Value.ToString();
        }
        #endregion

        #region int boxing operators
        public static implicit operator Row(int value)
        {
            return Row.getRow(value);
        }

        public static implicit operator int(Row row)
        {
            return row.Value;
        }

        public static bool operator ==(Row row, int intv)
        {
            return row.Value == intv;
        }

        public static bool operator !=(Row row, int intv)
        {
            return row.Value != intv;
        }

        public static bool operator ==(int intv, Row row)
        {
            return row.Value == intv;
        }

        public static bool operator !=(int intv, Row row)
        {
            return row.Value != intv;
        }
        #endregion

    }
}
