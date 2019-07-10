using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Api.Globalization
{
    public interface ILocalization
    {
        object ToLocal(object value);
    }
}
