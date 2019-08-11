using System;
using System.Collections.Generic;
using Kipon.Excel.Linq;

namespace Kipon.Excel.Cmd.Export
{
    public class ExportTest
    {
        public static void Run()
        {
            var data = new Data[]
            {
                new Data{ Id = Guid.NewGuid(), Name = "Name 1", Amount = 10.4M, Count = 10, Date = DateTime.Now, Has = true },
                new Data{ Id = Guid.NewGuid(), Name = "Name 2", Amount = 11.4M, Count = 10, Date = DateTime.Today, Has = false },
                new Data{ Id = Guid.NewGuid(), Name = "Name 3", Amount = 11.4M, Count = 12, Date = DateTime.Today.AddDays(2), Has = true }
            };

            var sheet = new Export.Sheets();
            sheet.Sheet1 = data;

            var list = new List<Data>();
            list.AddRange(data);
            list.Add(new Data { Id = Guid.NewGuid(), Name = "Name 4", Amount = 14.67M, Count = 23, Date = DateTime.Now.AddMonths(-3), Has = true, SecondDate = System.DateTime.Now });

            sheet.Sheet2 = list.ToArray();

            var excel = sheet.ToExcel();
            System.IO.File.WriteAllBytes(@"C:\Temp\excel-test.xlsx", excel);

        }
    }
}
