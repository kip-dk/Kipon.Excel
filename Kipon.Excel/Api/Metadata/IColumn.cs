using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Api.Metadata
{
    public interface IColumn
    {
        Api.Types.Column Min { get; set; }
        Api.Types.Column Max { get; set; }
        double Width { get; }
        bool Hidden { get; }
    }
}
