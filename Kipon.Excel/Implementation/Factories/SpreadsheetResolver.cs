using Kipon.Excel.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    public class SpreadsheetResolver<T, I> : BaseResolver<Kipon.Excel.Api.ISpreadsheet, I> 
    {
        public override Kipon.Excel.Api.ISpreadsheet Resolve(I instance)
        {
            var def = base.Resolve(instance);
            if (def != null)
            {
                return def;
            }

            var sheetResolver = new SheetResolver<IEnumerable<ISheet>, I, Models.Sheets<I>>();

            var sheets = sheetResolver.Resolve(instance);
            var spreadsheet = new Models.Spreadsheet(sheets);

            return spreadsheet;
        }
    }
}
