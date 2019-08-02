using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.ReaderImplementation.Models
{
    internal class Sheet : Kipon.Excel.Api.ISheet
    {
        private List<Cell> cells = new List<Cell>();


        public string Title { get; set; }

        public IEnumerable<ICell> Cells => cells;

        internal void Add(int zerobasecolumnindex, int zerobasedrowindex, object value)
        {
            cells.Add(new Cell(zerobasecolumnindex, zerobasedrowindex, value));
        }
    }
}
