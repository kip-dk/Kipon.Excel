using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Models.Cell
{
    internal class Cell : AbstractBaseCell, Kipon.Excel.Api.Cell.IDataType
    {
        internal Cell(int zbColumn, int zbRow, object value) : base(zbColumn, zbRow)
        {
            this.value = value;
            if (value != null)
            {
                this.ValueType = value.GetType();
            } 
        }

        private object value;

        public override object Value => this.value;

        public Type ValueType { get; set; }
    }
}
