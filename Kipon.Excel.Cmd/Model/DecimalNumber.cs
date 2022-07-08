using System;
using System.Collections.Generic;
using System.Text;
using Kipon.Excel.Attributes;

namespace Kipon.Excel.Cmd.Model
{
    public class DecimalNumber
    {
        [Sort(1)]
        public string Name { get; set; }

        [Sort(2)]
        [Decimals(0)]
        public decimal Value0 { get; set; }


        [Sort(3)]
        [Decimals(1)]
        public decimal Value1 { get; set; }

        [Sort(4)]
        [Decimals(2)]
        public decimal Value2 { get; set; }

        [Sort(5)]
        [Decimals(3)]
        public decimal Value3 { get; set; }

        [Sort(5)]
        [Decimals(4)]
        public decimal Value4 { get; set; }

        [Sort(5)]
        [Decimals(5)]
        public decimal Value5 { get; set; }

    }
}
