using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Models.Cell
{
    internal class Cell : AbstractBaseCell
    {
        internal Cell(int zbColumn, int zbRow, object value) : base(zbColumn, zbRow)
        {
            this.value = value;
        }

        private object value;

        public override object Value => this.value;
    }
}
