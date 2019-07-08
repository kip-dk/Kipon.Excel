using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel
{
    public interface ISheet
    {
        string Title { get; }
        ISheetProperties SheetProperties { get; }
        IEnumerable<IColumn> Columns { get; }
        IEnumerable<IRow> Rows { get; }
    }
}
