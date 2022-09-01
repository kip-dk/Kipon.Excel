using System;
using System.Collections.Generic;
using System.Text;
using Kipon.Excel.Attributes;

namespace Kipon.Excel.Cmd.Model
{
    public class MultiIndexSheet
    {
        [Sort(1)]
        public Index[] Sheet1 { get; set; }

        [Sort(2)]
        public Index[] Sheet2 { get; set; }

        public class Index
        {
            [Title("Name")]
            [Sort(1)]
            public string Name { get; set; }

            [Sort(2)]
            [Decimals(2)]
            public decimal Explicit { get; set; }

            [Sort(3)]
            public decimal Implicit { get; set; }

            [IndexColumn("[ab]-")]
            [Sort(3)]
            public Dictionary<string, int> Counts { get; set; } = new Dictionary<string, int>(); 

            [Ignore]
            public int this[string index]
            {
                get
                {
                    if (this.Counts.TryGetValue(index, out int v))
                    {
                        return v;
                    }
                    return 0;
                } set
                {
                    this.Counts[index] = value;
                }
            }
        }
    }
}
