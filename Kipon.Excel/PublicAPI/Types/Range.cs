using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Types
{
    /// <summary>
    /// A range represent a range in the matrix, where From is always the upper left corner and To is always the lower right corner.
    /// The consequence of this rule, is that the cells parsed into the constructor of range is not nesserarrly the same as thouse return
    /// from the From and To properties
    /// </summary>
    public struct Range
    {
        #region fields
        private Cell _from;
        private Cell _to;
        private string _value;
        #endregion

        #region constructors
        /// <summary>
        /// Create a range from two cell definitions.
        /// Be aware that cell1 do not nessesarrly become From, hence cell to might not become To
        /// </summary>
        /// <param name="cell1">The first cell in the range</param>
        /// <param name="cell2">The second cell in the range</param>
        public Range(Cell cell1, Cell cell2)
        {
            this._from = cell1.ToTopLeftRangeCorner(cell2);
            this._to = cell2.ToBottomRightRangeCorner(cell1);
            this._value = $"{_from.Value.ToString()}:{_to.Value.ToString()}";
        }


        /// <summary>
        /// It is valid to input ranges in wrong order, but the resulting range will be from left upper corner to right lower corner
        /// hence,  C2:A1 is valid, but will end up with a range defined as A1:C2
        /// </summary>
        /// <param name="value">A range string on from A1:C2</param>
        public Range(string value)
        {
            if (value == null) throw new NullReferenceException("value cannot be null");
            var spl = value.Split(':');
            if (spl.Length != 2) throw new FormatException("value is not in correct format, expected two cell reference on form A1:B2");
            var c1 = new Cell(spl[0]);
            var c2 = new Cell(spl[1]);
            this._from = c1.ToTopLeftRangeCorner(c2);
            this._to = c2.ToBottomRightRangeCorner(c1);
            this._value = $"{_from.Value.ToString()}:{_to.Value.ToString()}";
        }
        #endregion

        #region properties
        /// <summary>
        /// The upper left corner of the range
        /// </summary>
        public Cell From
        {
            get
            {
                return this._from;
            }
        }

        /// <summary>
        /// The lower right corner of the range
        /// </summary>
        public Cell To
        {
            get
            {
                return this._to;
            }
        }

        /// <summary>
        /// The natural string representation of the range, ex A1:B2, always in the order opperleftcorner:lowerrightcorner
        /// </summary>
        public string Value
        {
            get
            {
                return this._value;
            }
        }
        #endregion

        #region autoboxing
        public static implicit operator Range(string v)
        {
            return new Range(v);
        }

        public static implicit operator string(Range v)
        {
            return v.Value;
        }

        public static bool operator ==(Range v, string s)
        {
            return v.Value == s;
        }

        public static bool operator !=(Range v, string s)
        {
            return v.Value != s;
        }

        public static bool operator ==(string s, Range v)
        {
            return s == v.Value;
        }

        public static bool operator !=(string s, Range v)
        {
            return s != v.Value;
        }
        #endregion

        #region overrides
        public override bool Equals(object obj)
        {
            if (obj is Range)
            {
                return this.Value == ((Range)obj).Value;
            }

            if (obj is string)
            {
                return this.Value == (string)obj;
            }
            return false;
        }
        public override string ToString()
        {
            return this._value;
        }

        public override int GetHashCode()
        {
            return this._value.GetHashCode();
        }
        #endregion
    }
}
