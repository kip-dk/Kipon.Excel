using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel
{
    public interface IColumn
    {
        Types.Column Min { get; set; }
        Types.Column Max { get; set; }
        double Width { get; }
        bool Hidden { get; }
    }
}
