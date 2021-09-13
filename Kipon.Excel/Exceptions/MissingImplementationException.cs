using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Exceptions
{
    public class MissingImplementationException : ExcelBaseException
    {
        public MissingImplementationException(Type type) : base($"The sheet was expected to implement { type.FullName }.")
        {
        }
    }
}
