using Kipon.Excel.Attributes;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Linq;

namespace Kipon.Excel.UnitTests.TempTests
{
    public class HyperlineTest
    {
        [Test]
        public void GenHyperLinkSheet()
        {
            var cs = new UrlTest[]
            {
                new UrlTest{ Felt = "Test", Url = "https://kipon.dk;en tekst 1" },
                new UrlTest{ Felt = "Test", Url = "https://kipon.dk;en tekst 2" }
            };

            var data = cs.ToExcel();
            System.IO.File.WriteAllBytes(@"C:\Temp\testmedurl.xlsx", data);
        }


        public class UrlTest
        {
            [Title("Test")]
            public string Felt { get; set; }

            [Title("Url")]
            public string Url { get; set; }
        }


    }
}
