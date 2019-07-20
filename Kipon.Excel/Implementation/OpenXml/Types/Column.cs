using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.OpenXml.Types
{
    /// <summary>
    /// A column is a strong type for handling excel columns. It can be set from both int and strings
    /// It validates that a column name is valid (conforms to both naming and max columns according to excel spec), 
    /// and convert between the 0 based column index and its
    /// corresponding name (A,B,AA, DFE etc.)
    /// Created from string, it will calculate the corresponding 0 based index, created from int, it will calculate
    /// the corresponding string name, both exposed as Value and Index, - both immutable
    /// </summary>
    public struct Column
    {
        #region fields
        public const int EXCEL_TOTAL_MAXCOLUMNS = 16384;
        private const string valid_letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        private string _value;
        private int _index;
        #endregion

        #region constructors
        public Column(int no)
        {
            if (no < 0) throw new ArgumentOutOfRangeException("no cannot be less than 0");
            if (no >= EXCEL_TOTAL_MAXCOLUMNS) throw new ArgumentOutOfRangeException($"no cannot be greater or equal {EXCEL_TOTAL_MAXCOLUMNS}");

            this._value = Column.GetExcelColumnName(no + 1);
            this._index = no;
        }

        public Column(string v)
        {
            if (v == null) throw new NullReferenceException("v cannot be null");
            if (v.Length < 1) throw new ArgumentException("length of v must be at least 1");
            if (v.Length > 3) throw new ArgumentException("length of v must be at most 3");

            foreach (var c in v.Distinct())
            {
                if (!valid_letters.Contains(c))
                {
                    throw new ArgumentException($"{c} is not a valid character for a column name");
                }
            }

            _value = v;
            this._index = ColumnNameToNumber(v);

            if (this._index >= EXCEL_TOTAL_MAXCOLUMNS) throw new ArgumentOutOfRangeException($"no cannot be greater or equal {EXCEL_TOTAL_MAXCOLUMNS}");

        }
        #endregion

        #region int boxing operators
        public static implicit operator Column(int v)
        {
            return new Column(v);
        }
        public static implicit operator int(Column v)
        {
            return v.Index;
        }

        public static bool operator ==(Column v, int i)
        {
            return v.Index == i;
        }

        public static bool operator !=(Column v, int i)
        {
            return v.Index != i;
        }

        public static bool operator ==(int i, Column v)
        {
            return i == v.Index;
        }

        public static bool operator !=(int i, Column v)
        {
            return i != v.Index;
        }
        #endregion

        #region string boxing
        public static implicit operator Column(string v)
        {
            return new Column(v);
        }

        public static implicit operator string(Column v)
        {
            return v.Value;
        }

        public static bool operator ==(Column v, string s)
        {
            return v.Value == s;
        }

        public static bool operator != (Column v, string s)
        {
            return v.Value != s;
        }

        public static bool operator ==(string s, Column v)
        {
            return s == v.Value;
        }

        public static bool operator !=(string s, Column v)
        {
            return s != v.Value;
        }
        #endregion

        #region overrides
        public override string ToString()
        {
            return this.Value;
        }

        public override int GetHashCode()
        {
            return this.Value.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (obj is Column)
            {
                return this.Value == ((Column)obj).Value;
            }

            if (obj is string)
            {
                return this.Value == (string)obj;
            }

            if (obj is int)
            {
                return this.Index == (int)obj;
            }

            if (obj is int? && obj != null)
            {
                return this.Index == ((int?)obj).Value;
            }
            return false;
        }
        #endregion

        #region properties
        public string Value
        {
            get
            {
                return this._value;
            }
        }

        public int Index
        {
            get
            {
                return this._index;
            }
        }
        #endregion

        #region helper converter methods
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private static int ColumnNameToNumber(string col_name)
        {
            int result = 0;

            // Process each letter.
            for (int i = 0; i < col_name.Length; i++)
            {
                result *= 26;
                char letter = col_name[i];

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                result += (int)letter - (int)'A' + 1;
            }
            return result - 1;
        }
        #endregion
    }
}
