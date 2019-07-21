using Kipon.Excel.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Models.Sheets
{
    internal class ArrayOfISheet<I> : AbstractBaseSheets<I>
    {
        public override void Populate(I instance)
        {
            var sheets = instance as ISheet[];
            foreach (var sheet in sheets)
            {
                this.Add(sheet);
            }
        }
    }
}
