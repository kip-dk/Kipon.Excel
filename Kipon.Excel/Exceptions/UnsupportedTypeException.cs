using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Exceptions
{
    public class UnsupportedTypeException : ExcelBaseException
    {
        public UnsupportedTypeException(Type t) : base("Unsupported type " + t.FullName, null)
        {
        }

        public UnsupportedTypeException(string methodname, Type t) : base($"{methodname} does not support type " + t.FullName, null)
        {
        }

        public UnsupportedTypeException(string methodname, Type t, string suggestion) : base($"{methodname} does not support type {t.FullName}. {suggestion}", null)
        {
        }
    }
}
