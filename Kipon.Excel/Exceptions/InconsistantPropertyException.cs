using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Exceptions
{
    public sealed class InconsistantPropertyException : ExcelBaseException
    {
        public InconsistantPropertyException(System.Reflection.PropertyInfo property) : base($"Property {property.Name} has inconsistant definition. getter method might be missing, or IgnoreAttribute flag has been set together with ColumnAttribute flag.")
        {
        }
    }
}
