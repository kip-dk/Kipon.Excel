using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]

    public class MaxlengthAttribute : Attribute
    {
        private int _max;

        public MaxlengthAttribute(int maxlen)
        {
            if (maxlen <= 0) throw new ArgumentOutOfRangeException("maxlen cannot be less than or 0");
            this._max = maxlen;
        }

        public int Max
        {
            get
            {
                return this._max;
            }
        }
    }
}
