using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Serialization
{
    internal interface IStyleResolver
    {
        uint Resolve(ISheet sheet, Types.Cell cell);
    }
}
