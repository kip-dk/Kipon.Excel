using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Attributes
{
    /// <summary>
    /// Use the Sheet Attribute to decorate every property in your "sheets" class to indicate that the
    /// property is representing a sheet. If at most one property is decorated with sheet in your class,
    /// the class will be considered a "sheet-attribute-decorated" sheets list, and only properties with this attribute will be
    /// streamed. Inherited properties with the attribute will be streamed as well. Only public instance properties
    /// will be included.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class SheetAttribute : Attribute
    {
        public string Title { get; set; }
        public int? Sort { get; set; }

        public SheetAttribute()
        {
        }

        public SheetAttribute(string title)
        {
            this.Title = title;
        }

        public SheetAttribute(int sort)
        {
            this.Sort = sort;
        }

        public SheetAttribute(string title, int sort)
        {
            this.Title = title;
            this.Sort = sort;
        }
    }
}
