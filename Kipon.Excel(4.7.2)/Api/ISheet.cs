using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Api
{
    public interface ISheet
    {
        string Title { get; }
        IEnumerable<ICell> Cells { get; }
    }
}
