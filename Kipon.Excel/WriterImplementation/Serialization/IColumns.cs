using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.WriterImplementation.Serialization
{
    internal interface IColumns
    {
        IEnumerable<IColumn> Columns { get; }
    }
}
