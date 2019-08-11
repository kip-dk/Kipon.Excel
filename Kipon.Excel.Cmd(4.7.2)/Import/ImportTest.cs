using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Linq;

namespace Kipon.Excel.Cmd.Import
{
    public class ImportTest
    {
        public static void Run()
        {
            using (var fs = new System.IO.FileStream(@"C:\Temp\excel-test.xlsx", System.IO.FileMode.Open))
            {
                var sp = fs.ToObject<Kipon.Excel.Cmd.Export.Sheets>();
            }
        }
    }
}
