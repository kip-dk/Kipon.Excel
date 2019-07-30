using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.WriterImplementation.Models
{
    internal class Spreadsheet : ISpreadsheet
    {
        private IEnumerable<ISheet> _sheets;

        internal Spreadsheet(IEnumerable<ISheet> sheets)
        {
            if (sheets == null) throw new ArgumentNullException("sheets cannot be null");
            if (sheets.Count() < 1) throw new ArgumentOutOfRangeException("sheets must have a least one instance");
            this._sheets = sheets;
        }


        public IEnumerable<ISheet> Sheets
        {
            get
            {
                return this._sheets;
            }
        }
    }
}
