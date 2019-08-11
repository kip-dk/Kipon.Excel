using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Exceptions
{
    public class NullInstanceException : ExcelBaseException
    {
        public NullInstanceException(Type t) : base($"{t.FullName} cannot be resolved from a null value")
        {
        }
    }
}
