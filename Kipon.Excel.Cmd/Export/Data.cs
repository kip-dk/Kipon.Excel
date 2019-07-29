using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Cmd.Export
{
    public class Data
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public decimal Amount { get; set; }
        public int Count { get; set; }
        public bool Has { get; set; }
        public DateTime Date { get; set; }
    }
}
