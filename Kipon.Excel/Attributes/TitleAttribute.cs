using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class TitleAttribute : Attribute
    {
        private string _value;
        public TitleAttribute(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                throw new ArgumentException("value cannot be null or empty");
            }

            this._value = value;
        }

        public string Value => this._value;
    }
}
