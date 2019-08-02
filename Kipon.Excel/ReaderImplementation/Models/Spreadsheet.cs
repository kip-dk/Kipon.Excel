using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.ReaderImplementation.Models
{
    internal class Spreadsheet : Kipon.Excel.Api.ISpreadsheet
    {
        private List<Sheet> sheets = new List<Sheet>();

        public IEnumerable<ISheet> Sheets => this.sheets;


        internal Sheet Add(string title)
        {
            var result = new Sheet { Title = title };
            sheets.Add(result);
            return result;
        }
    }
}
