using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.Implementation.Models
{
    internal class Sheet : ISheet
    {
        private IEnumerable<ICell> _cell;
        public Sheet()
        {

        }

        public string Title { get; set; }

        public IEnumerable<ICell> Cells
        {
            get
            {
                if (_cell == null)
                {
                    return new ICell[0];
                }
                return _cell;
            }
            set
            {
                if (value == null) throw new NullReferenceException("value og Cells cannot be null");
                this._cell = value;
            }
        }
    }
}
