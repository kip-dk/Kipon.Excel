using System;
using System.Collections.Generic;
using Kipon.Excel.Linq;

namespace Kipon.Excel.Cmd
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = new List<Model.DecimalNumber>();

            data.Add(new Model.DecimalNumber { Name = "r1", Value0 = 10M, Value1 = 1.1M, Value2 = 2.2M, Value3 = 3.3M, Value4 = 4.4M, Value5 = 5.5M });
            data.Add(new Model.DecimalNumber { Name = "r2", Value0 = 10M, Value1 = 10.1M, Value2 = 20.2M, Value3 = 30.3M, Value4 = 40.4M, Value5 = 50.5M });

            var excel = data.ToExcel();
            System.IO.File.WriteAllBytes(@"C:\Temp\decimal-test.xlsx", excel);

        }
    }
}
