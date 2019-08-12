using System;
using System.Collections.Generic;
using System.Text;
using Kipon.Excel.Attributes;
using Kipon.Excel.Linq;

namespace Kipon.Excel.Cmd.Model
{
    public class Employee
    {
        public static Employee[] FromExcel()
        {
            using (var fs = new System.IO.FileStream(@"C:\Temp\Tirsdag den 13.8 kl 18.15 – alle i vest.xlsx", System.IO.FileMode.Open))
            {
                return fs.ToArray<Employee>();
            }
        }

        [Title("Navn")]
        public string Name { get; set; }


        [Title("Mobil")]
        public string Mobile { get; set; }
    }
}
