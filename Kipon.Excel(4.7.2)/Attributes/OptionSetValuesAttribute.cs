using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class OptionSetValuesAttribute : Attribute
    {
        private string[] _values;

        public OptionSetValuesAttribute(string[] value)
        {
            this._values = value;
        }

        public string[] Value
        {
            get
            {
                return this._values;
            }
        }
    }
}
