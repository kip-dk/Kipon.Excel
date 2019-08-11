using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Api.Cell
{
    public interface IReadonly
    {
        bool IsReadonly { get; }
    }
}
