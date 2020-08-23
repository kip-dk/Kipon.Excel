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
        public int Sort { get; set; } = int.MinValue;
        public bool IsHidden { get; set; }
        public bool IsReadonly { get; set; }
        public int? Decimals { get; set; }
        public int? MaxLength { get; set; }
        public double? Width { get; set; }
        public  bool Optional { get; set; }
        public string[] OptionSetValues { get; set; }
        public ColumnAttribute() { }
    }
}
