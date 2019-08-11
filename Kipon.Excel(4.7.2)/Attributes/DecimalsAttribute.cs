using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class DecimalsAttribute : Attribute
    {
        public DecimalsAttribute(int decimals)
        {
            this.Value = decimals;
        }

        public int Value { get; private set; }
    }
}
