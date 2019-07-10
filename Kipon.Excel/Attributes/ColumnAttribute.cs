using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ColumnAttribute : Attribute
    {
        public string Title { get; set; }
        public int Sort { get; set; } = 0;
        public bool Hidden { get; set; }

        public ColumnAttribute() { }

        public ColumnAttribute(string title, int sort)
        {
            this.Title = title;
            this.Sort = sort;

        }
    }
}
