using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class IndexColumnAttribute : Attribute
    {
        public string Pattern { get; private set; }

        public System.Text.RegularExpressions.Regex Expression { get; private set; }

        public IndexColumnAttribute(string pattern) 
        {
            if (string.IsNullOrEmpty(pattern))
            {
                throw new ArgumentException(pattern);
            }
            this.Pattern = pattern;

            try
            {
                this.Expression = new System.Text.RegularExpressions.Regex(this.Pattern);
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Cannot parse index column pattern to regular expression {pattern}: {ex.Message}");
            }
        }
    }
}
