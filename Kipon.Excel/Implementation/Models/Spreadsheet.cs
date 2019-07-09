using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Models
{
    internal class Spreadsheet : Kipon.Excel.ISpreadsheet
    {
        private List<ISheet> _sheets;

        public IEnumerable<ISheet> Sheets
        {
            get
            {
                if (_sheets == null)
                {
                    return new ISheet[0];
                }

                return this._sheets;
            }
            set
            {
                if (value == null) throw new NullReferenceException("value of property Sheets cannot be set to null");
                this._sheets = value.ToList();
            }
        }
    }
}
