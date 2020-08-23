using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Extensions.Strings
{
    internal static class StringMethods
    {
        internal static string ToRelaxedName(this string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            return value.Replace(" ", "").Replace("\t","").Replace("\r","").Replace("\n","").Trim().ToLower();
        }
    }
}
