using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.ReaderImplementation.Models
{
    internal class Cell : Kipon.Excel.Api.ICell
    {
        private Coordinate coordinate;

        public Cell(int zerobasedcolumnindex, int zerobasedrowindex, object o)
        {
            this.coordinate = new Coordinate(zerobasedcolumnindex, zerobasedrowindex);
            this.Value = o;
        }


        public ICoordinate Coordinate => coordinate;

        public object Value { get; private set; }
    }
}
