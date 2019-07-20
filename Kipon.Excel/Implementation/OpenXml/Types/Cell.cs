using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.OpenXml.Types
{
    /// <summary>
    /// A cell represent a reference to a cell in a sheet by its name. It is by nature build from a column and a row reference, and as such it
    /// simply points into coordinate of the sheet.
    /// The upper left value of a cell is A1, corresponding to a column A (index=0), and a row value 1, (index 0). Hince, a cell has it natual name.
    /// </summary>
    public struct Cell
    {
        #region fields
        private Column _column;
        private Row _row;
        private string _value;
        #endregion

        #region constructors
        /// <summary>
        /// Creates a Cell reference from an instance of column and row
        /// </summary>
        /// <param name="column">column name, ex. A</param>
        /// <param name="row">ZERO based row, ex 0</param>
        private static Cell CellFromCR(Column column, Row row)
        {
            if (column == null) throw new NullReferenceException("column cannot be null");

            return column.Value + (row.Value + 1).ToString();
        }

        /// <summary>
        /// Create a Cell reference from its natural name
        /// </summary>
        /// <param name="value">Natural name of cell,  ex A1 for the first cell in the sheet</param>
        private Cell(string value)
        {
            if (value == null) throw new NullReferenceException("value cannot be null");
            if (value.Length < 2) throw new ArgumentException("length of value cannot be less than 2");

            var columnName = new System.Text.StringBuilder();
            var ix = 0;
            while (!Char.IsNumber(value[ix]))
            {
                columnName.Append(value[ix]);
                ix++;
            }

            if (columnName.Length < 1) throw new ArgumentException("value must start with valid column ref, ex A or BA etc.");

            var rowNumber = new System.Text.StringBuilder();
            while (ix < value.Length)
            {
                rowNumber.Append(value[ix]);
                ix++;
            }

            int lineNo = 0;
            if (!Int32.TryParse(rowNumber.ToString(), out lineNo))
            {
                throw new ArgumentException("value must contain a valid row reference after the column name");
            }

            this._column = new Column(columnName.ToString());
            this._row = new Row(lineNo - 1);
            this._value = value;
        }

#warning dirty method due to wrong ISheet api
        internal Cell(uint column, uint row)
        {
            this._column = new Column((int)column);
            this._row = new Row((int)row);
            _value = _column.Value + (_row.Value + 1).ToString();
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
            if (obj is Cell)
            {
                return this.Value == ((Cell)obj).Value;
            }

            if (obj is string)
            {
                return this.Value == (string)Value;
            }

            return false;
        }
        #endregion

        #region operators
#warning caching directly here is not pretty, we need a garbage collection mechanims to cleanup when no excel writing is going one
        private static readonly Dictionary<string, Cell> cells = new Dictionary<string, Cell>();
        public static implicit operator Cell(string v)
        {
            if (v == null) throw new NullReferenceException();
            lock (cells)
            {
                if (cells.ContainsKey(v)) return cells[v];
                cells[v] = new Cell(v);
                return cells[v];
            }
        }

        public static implicit operator string(Cell v)
        {
            return v.Value;
        }

        public static bool operator ==(Cell v, string s)
        {
            return v.Value == s;
        }

        public static bool operator !=(Cell v, string s)
        {
            return s != v.Value;
        }

        public static bool operator ==(string s, Cell v)
        {
            return s == v.Value;
        }

        public static bool operator !=(string s, Cell v)
        {
            return s != v.Value;
        }
        #endregion

        #region properties
        public Column Column
        {
            get
            {
                return _column;
            }
        }

        public Row Row
        {
            get
            {
                return _row;
            }
        }

        public string Value
        {
            get
            {
                return _value;
            }
        }
        #endregion

        #region internal helper methods
        internal Cell ToTopLeftRangeCorner(Cell other)
        {
            if (this.Value == other.Value) return this;

            int column = this.Column.Index;
            if (column > other.Column.Index)
            {
                column = other.Column.Index;
            }

            int row = this.Row.Value;
            if (row > other.Row.Value)
            {
                row = other.Row.Value;
            }

            return CellFromCR(column, row);
        }

        internal Cell ToBottomRightRangeCorner(Cell other)
        {
            if (this.Value == other.Value) return this;

            int column = this.Column.Index;
            if (column < other.Column.Index)
            {
                column = other.Column.Index;
            }

            int row = this.Row.Value;
            if (row < other.Row.Value)
            {
                row = other.Row.Value;
            }

            return CellFromCR(column, row);
        }
        #endregion
    }
}
