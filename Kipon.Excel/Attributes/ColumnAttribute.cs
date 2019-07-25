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
        public int? Sort { get; set; } 
        public bool IsHidden { get; set; }
        public bool IsReadonly { get; set; }
        public ColumnAttribute() { }

        public ColumnAttribute(string title)
        {
            this.Title = title;
        }

        public ColumnAttribute(int sort)
        {
            this.Sort = sort;
        }

        public ColumnAttribute(string title, int sort)
        {
            this.Title = title;
            this.Sort = sort;

        }
    }
}
