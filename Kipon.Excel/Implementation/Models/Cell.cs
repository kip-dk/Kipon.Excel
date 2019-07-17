using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.Implementation.Models
{
    internal class Cell : ICell
    {

        public Kipon.Excel.Api.Types.Cell Name { get; set; }

        public object Value
        {
            get; set;
        }
    }
}
