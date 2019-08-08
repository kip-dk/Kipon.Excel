using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Exceptions
{
    public class DefaultConstructorRequiredException : ExcelBaseException
    {
        public DefaultConstructorRequiredException(Type t) : base($"Type {t.FullName} does not have the required default constructor.", null)
        {
        }
    }
}
