using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Exceptions
{
    public class UnresolveableTypeException : ExcelBaseException
    {
        public UnresolveableTypeException(Type resolveFrom, Type resolveTo) : base($"Type {resolveFrom.FullName} cannot be resolved to Type {resolveTo.FullName}")
        {
        }
    }
}
