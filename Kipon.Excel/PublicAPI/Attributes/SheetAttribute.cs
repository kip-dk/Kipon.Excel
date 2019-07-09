using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.PublicAPI.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class SheetAttribute : Attribute
    {
        public string Title { get; set; }
        public int Sort { get; set; }

        public SheetAttribute()
        {
        }

        public SheetAttribute(string title, int sort)
        {
            this.Title = title;
            this.Sort = sort;
        }
    }
}
