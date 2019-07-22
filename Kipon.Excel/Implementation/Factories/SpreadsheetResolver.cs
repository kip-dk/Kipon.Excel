using Kipon.Excel.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    internal class SpreadsheetResolver : AbstractBaseResolver<Kipon.Excel.Api.ISpreadsheet> 
    {
        public override Kipon.Excel.Api.ISpreadsheet Resolve(object instance)
        {
            var def = base.Resolve(instance);
            if (def != null)
            {
                return def;
            }

            var sheetResolver = new SheetsResolver<Models.Sheets.AbstractBaseSheets>();

            var sheets = sheetResolver.Resolve(instance);
            var spreadsheet = new Models.Spreadsheet(sheets);

            return spreadsheet;
        }
    }
}
