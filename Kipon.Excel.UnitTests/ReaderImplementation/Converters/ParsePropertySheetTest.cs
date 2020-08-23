using NUnit.Framework;
using Kipon.Excel.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.UnitTests.Fake.Data;
using Kipon.Excel.Attributes;
using Kipon.Excel.Linq;

namespace Kipon.Excel.UnitTests.ReaderImplementation.Converters
{
    public class ParsePropertySheetTest
    {
        [Test]
        public void ConverterTest()
        {
            using (var excel = this.GetType().Assembly.GetManifestResourceStream("Kipon.Excel.UnitTests.Resources.ParsePropertySheetTest.xlsx"))
            {
                var result = excel.ToArray<ParsePropertySheet>();
                Assert.AreEqual(3, result.Length);

                Assert.AreEqual(1, result[0].FirstField);
                Assert.AreEqual(2, result[1].FirstField);
                Assert.AreEqual(3, result[2].FirstField);

                Assert.AreEqual("a", result[0].Nextfield);
                Assert.AreEqual("b", result[1].Nextfield);
                Assert.AreEqual("c", result[2].Nextfield);

                Assert.AreEqual("1a", result[0].Another);
                Assert.AreEqual("2b", result[1].Another);
                Assert.AreEqual("3c", result[2].Another);

            }
        }

        [Parse(FirstColumn = 2, FirstRow = 2)]
        public class ParsePropertySheet
        {
            public int FirstField { get; set; }
            public string Nextfield { get; set; }
            public string Another { get; set; }
        }
    }
}
