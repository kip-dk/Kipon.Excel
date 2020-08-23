using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Attributes
{
    /// <summary>
    /// The parse attribute can give the excel parser special instructions on parsing the the data.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ParseAttribute : Attribute
    {
        /// <summary>
        /// Tell the parser to read data in another start row than the first
        /// </summary>
        public int FirstRow { get; set; } = 1;

        /// <summary>
        /// Tell the parser to read data in another start column than the first
        /// </summary>
        public int FirstColumn { get; set; } = 1;

        public ParseAttribute() { }

        public ParseAttribute(int firstRow, int firstColumn) 
        {
            this.FirstRow = firstRow;
            this.FirstColumn = FirstColumn;
        }
    }
}
