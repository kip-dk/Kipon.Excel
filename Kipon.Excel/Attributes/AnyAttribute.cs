using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    internal class AnyAttribute : Attribute
    {
    }
}
