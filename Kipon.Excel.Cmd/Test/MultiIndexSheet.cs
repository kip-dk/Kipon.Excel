using System;
using System.Collections.Generic;
using System.Text;
using Kipon.Excel.Linq;

namespace Kipon.Excel.Cmd.Test
{
    public class MultiIndexSheet
    {
        public void Execute()
        {
            var model = new Model.MultiIndexSheet
            {
                Sheet1 = new Model.MultiIndexSheet.Index[]
                {
                    new Model.MultiIndexSheet.Index{ Name = "index 1 - Row 1", ["a-1"] = 1, ["a-2"] = 2 },
                },
                Sheet2 = new Model.MultiIndexSheet.Index[]
                {
                    new Model.MultiIndexSheet.Index{ Name = "index 2 - Row 1", ["b-1"] = 1, ["b-2"] = 2 }
                }
            };

            var excel = model.ToExcel();

            System.IO.File.WriteAllBytes(@"C:\Temp\multitest.xlsx", excel);
        }
    }
}
