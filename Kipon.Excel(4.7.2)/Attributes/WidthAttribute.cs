using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class WidthAttribute : Attribute
    {
        private double _value;
        public WidthAttribute(double value)
        {
            this._value = value;
        }

        public double Value
        {
            get
            {
                return this._value;
            }
        }
    }
}
