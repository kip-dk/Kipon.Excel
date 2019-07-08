using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel
{
    public interface IRange
    {
        int FromRow { get; }

        int ToRow { get; }
    }
}
