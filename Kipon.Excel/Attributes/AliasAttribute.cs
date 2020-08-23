using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public class AliasAttribute : Attribute
    {
        public string Title { get; set; }

        public AliasAttribute()
        {

        }

        public AliasAttribute(string title)
        {
            this.Title = title;
        }
    }
}
