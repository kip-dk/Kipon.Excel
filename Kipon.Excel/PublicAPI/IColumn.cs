using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel
{
    public interface IColumn
    {
        int Min { get; set; }
        int Max { get; set; }
        double Width { get; }
        bool Hidden { get; }
    }
}
