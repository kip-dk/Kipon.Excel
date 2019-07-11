using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Exceptions
{
    public class ExcelBaseException : Exception
    {
        public ExcelBaseException(string message, Exception innnerException = null): base(message, innnerException)
        {
        }
    }
}
