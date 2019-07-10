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
        private List<IRow> _rows;
        public Sheet()
        {

        }

        public string Title { get; set; }

        public IEnumerable<IRow> Rows
        {
            get
            {
                if (_rows == null)
                {
                    return new IRow[0];
                }
                return _rows;
            }
            set
            {
                if (value == null) throw new NullReferenceException("value og Rows cannot be null");
                this._rows = value.ToList();
            }
        }
    }
}
