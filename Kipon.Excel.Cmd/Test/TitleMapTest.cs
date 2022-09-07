using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using Kipon.Excel.Attributes;
using Kipon.Excel.Linq;

namespace Kipon.Excel.Cmd.Test
{
    public class TitleMapTest
    {
        public void Execute()
        {
            var rows = new string[] { "a", "b" };

            var r = (from ro in rows
                     select new MySheet
                     {
                         FirstColumn = $"1: { ro }",
                         SecondColumn = $"2: { ro }"
                     }).ToArray();

            var sheets = new Sheets
            {
                Sheet1 = r,
                Sheet2 = r
            };

            var excel = sheets.ToExcel();
            System.IO.File.WriteAllBytes(@"C:\Temp\mapped-column.xlsx", excel);
        }
    }

    public class Sheets
    {
        [Title("Sheet 1")]
        [Sort(1)]
        public MySheet[] Sheet1 { get; set; }

        [Title("Sheet 2")]
        [Sort(2)]
        public MySheet[] Sheet2 { get; set; }
    }

    public class MySheet : Kipon.Excel.Api.ITitleMap
    {
        public string MapFromExcelValue(string value)
        {
            return value;
        }

        public string MapToExcelValue(string value)
        {
            return value.Replace("[X]", "2022");
        }

        [Sort(1)]
        [Title("First Column [X]")]
        public string FirstColumn { get; set; }

        [Sort(2)]
        [Title("Second Column [X]")]
        public string SecondColumn { get; set; }
    }
}
