using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Models
{
    internal class Row : IRow
    {
        private List<ICell> _cells;

        public IEnumerable<ICell> Cells
        {
            get
            {
                if (_cells == null)
                {
                    return new ICell[0];
                }
                return _cells;
            }
            set
            {
                if (value == null) throw new NullReferenceException("Cells of Row cannot be null");
                this._cells = value.ToList();
            }
        }
    }
}
