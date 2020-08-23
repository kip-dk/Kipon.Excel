using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Kipon.Excel.Exceptions
{
    public class UnsupportedColumnIndexPropertyException : ExcelBaseException
    {
        public UnsupportedColumnIndexPropertyException(Type type, PropertyInfo property) : base($"Unsupported column index type for: {type.FullName}, Property {property.Name}. Only Dictionary<string, simple type> is supported")
        {

        }
    }
}
