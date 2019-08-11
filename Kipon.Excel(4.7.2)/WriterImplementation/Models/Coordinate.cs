using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.WriterImplementation.Models
{
    internal class Coordinate : Kipon.Excel.Api.ICoordinate
    {
        private int[] points;
        internal Coordinate(int zerobasedColumnIndex, int zeroBasedRowIndex)
        {
            this.points = new int[]
            {
                zerobasedColumnIndex, zeroBasedRowIndex
            };
        }

        public IEnumerable<int> Point => this.points;
    }
}
