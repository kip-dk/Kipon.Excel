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
        public void ParseAsArrayTest()
        {
            using (var excel = this.GetType().Assembly.GetManifestResourceStream("Kipon.Excel.UnitTests.Resources.ParsePropertySheetTest.xlsx"))
            {
                var result = excel.ToArray<ParsePropertySheet>();
                Assert.AreEqual(3, result.Length);

                Assert.AreEqual(1, result[0].AField);
                Assert.AreEqual(2, result[1].AField);
                Assert.AreEqual(3, result[2].AField);

                Assert.AreEqual("a", result[0].Nextfield);
                Assert.AreEqual("b", result[1].Nextfield);
                Assert.AreEqual("c", result[2].Nextfield);

                Assert.AreEqual("1a", result[0].WrongAnother);
                Assert.AreEqual("2b", result[1].WrongAnother);
                Assert.AreEqual("3c", result[2].WrongAnother);

                Assert.AreEqual(60, result[0].Sum);
                Assert.AreEqual(130, result[1].Sum);
                Assert.AreEqual(180, result[2].Sum);

                Assert.AreEqual(50, result[1].Values["P123662"]);
            }
        }

        [Test]
        public void ParseAsObjectTest()
        {
            using (var excel = this.GetType().Assembly.GetManifestResourceStream("Kipon.Excel.UnitTests.Resources.ParsePropertySheetTest.xlsx"))
            {
                var r1 = excel.ToObject<PrasePropertySheetContainer>();

                var result = r1.Properties;
                Assert.NotNull(result);
                Assert.AreEqual(3, result.Length);

                Assert.AreEqual(1, result[0].AField);
                Assert.AreEqual(2, result[1].AField);
                Assert.AreEqual(3, result[2].AField);

                Assert.AreEqual("a", result[0].Nextfield);
                Assert.AreEqual("b", result[1].Nextfield);
                Assert.AreEqual("c", result[2].Nextfield);

                Assert.AreEqual("1a", result[0].WrongAnother);
                Assert.AreEqual("2b", result[1].WrongAnother);
                Assert.AreEqual("3c", result[2].WrongAnother);

                Assert.AreEqual(60, result[0].Sum);
                Assert.AreEqual(130, result[1].Sum);
                Assert.AreEqual(180, result[2].Sum);

                Assert.AreEqual(50, result[1].Values["P123662"]);

                Assert.AreEqual(3, result[0].Row);
                Assert.AreEqual(4, result[1].Row);
                Assert.AreEqual(5, result[2].Row);
            }
        }

        [Parse(FirstColumn = 2, FirstRow = 2)]
        public class ParsePropertySheet : Kipon.Excel.Api.IRowAware
        {
            [Title("First field")]
            public int AField { get; set; }
            public string Nextfield { get; set; }


            [Alias("Another")] 
            public string WrongAnother { get; set; }

            [IndexColumn("P[123456789][0123456789]{3,7}")]
            public Dictionary<string, int> Values { get; set; }

            [Optional]
            public string NotThereIsOk { get; set; }

            public int Sum
            {
                get
                {
                    if (this.Values != null && this.Values.Count > 0)
                    {
                        return this.Values.Values.Sum();
                    }
                    return 0;
                }
            }

            [Kipon.Excel.Attributes.Ignore]
            public int Row
            {
                get; set;
            }

            public void SetRowNo(int excelRowNumber)
            {
                this.Row = excelRowNumber;
            }
        }

        public class PrasePropertySheetContainer
        {
            public ParsePropertySheet[] Properties { get; set; }
        }
    }
}
