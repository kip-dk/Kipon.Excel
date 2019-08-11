using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Attributes;

namespace Kipon.Excel.Cmd.Export
{
    public class Data
    {
        [Hidden]
        [Sort(1)]
        public Guid Id { get; set; }
        [Sort(2)]
        public string Name { get; set; }
        [Sort(3)]
        public decimal Amount { get; set; }
        [Sort(4)]
        public int Count { get; set; }
        [Sort(5)]
        public bool Has { get; set; }
        [Sort(6)]
        public DateTime Date { get; set; }
        [Sort(7)]
        public DateTime? SecondDate { get; set; }
    }
}
