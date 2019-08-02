using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.ReaderImplementation.Models
{
    internal class Coordinate : Kipon.Excel.Api.ICoordinate
    {
        private int[] coordinates;

        internal Coordinate(int zerobasedcolumnindex, int zerobasedrowindex)
        {
            this.coordinates = new int[]
            {
                zerobasedcolumnindex,
                zerobasedrowindex
            };
        }

        public IEnumerable<int> Point => coordinates;
    }
}
