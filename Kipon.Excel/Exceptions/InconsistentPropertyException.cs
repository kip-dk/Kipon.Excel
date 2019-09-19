using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Exceptions
{
    public sealed class InconsistentPropertyException : ExcelBaseException
    {
        public InconsistentPropertyException(System.Reflection.PropertyInfo property) : base($"Property {property.Name} has inconsistent property definition. getter method might be missing for property with ColumnAttribute, or IgnoreAttribute flag has been set together with ColumnAttribute flag.")
        {
        }
    }
}
