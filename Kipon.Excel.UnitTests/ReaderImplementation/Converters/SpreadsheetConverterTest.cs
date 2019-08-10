using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Linq;


namespace Kipon.Excel.UnitTests.ReaderImplementation.Converters
{
    [TestClass]
    public class SpreadsheetConverterTest
    {
        private PropertySheets sheet;
        private byte[] sheetData;

        private Fake.Data.ValueProperty[] arrayValues;
        private byte[] arrayValuesData;

        private void Setup()
        {
            {
                this.sheet = new PropertySheets();
                var numbers1 = new short[] { 1, 2, 4 };

                sheet.S1 = (from number in numbers1
                            select new Fake.Data.ValueProperty(number)).ToArray();
            }

            {
                var numbers2 = new short[] { 7, 9, 12 };

                sheet.S2 = (from number in numbers2
                            select new Fake.Data.ValueProperty(number)).ToList();

                this.sheetData = this.sheet.ToExcel();
            }

            {
                var numbers3 = new short[] { 14, 16, 18, 21 };
                this.arrayValues = (from number in numbers3
                                    select new Fake.Data.ValueProperty(number)).ToArray();

                this.arrayValuesData = arrayValues.ToExcel();
            }
        }

        [TestMethod]
        public void ConvertPropertySheetTest()
        {
            this.Setup();
            using (var mem = new System.IO.MemoryStream(sheetData))
            {
                var converter = new Kipon.Excel.ReaderImplementation.Converters.SpreadsheetConverter();
                var result = converter.Convert<PropertySheets>(mem);

                var ix = 0;
                foreach (var vp in result.S1)
                {
                    Assert.AreEqual(sheet.S1[ix].Boolean, vp.Boolean);
                    Assert.AreEqual(sheet.S1[ix].BooleanNullable, vp.BooleanNullable);

                    Assert.AreEqual(sheet.S1[ix].DateTime, vp.DateTime);
                    Assert.AreEqual(sheet.S1[ix].DateTimeNullable, vp.DateTimeNullable);

                    Assert.AreEqual(sheet.S1[ix].Decimal, vp.Decimal);
                    Assert.AreEqual(sheet.S1[ix].DecimalNullable, vp.DecimalNullable);

                    Assert.AreEqual(sheet.S1[ix].Double, vp.Double);
                    Assert.AreEqual(sheet.S1[ix].DoubleNullable, vp.DoubleNullable);

                    Assert.AreEqual(sheet.S1[ix].Guid, vp.Guid);
                    Assert.AreEqual(sheet.S1[ix].GuidNullable, vp.GuidNullable);
                    ix++;
                }

                ix = 0;
                foreach (var vp in result.S2)
                {
                    Assert.AreEqual(sheet.S2[ix].Guid, vp.Guid);
                    Assert.AreEqual(sheet.S2[ix].GuidNullable, vp.GuidNullable);
                    ix++;
                }
            }
        }

        [TestMethod]
        public void ConvertArrayTest()
        {
            this.Setup();
            using (var mem = new System.IO.MemoryStream(arrayValuesData))
            {
                var converter = new Kipon.Excel.ReaderImplementation.Converters.SpreadsheetConverter();
                var result = converter.Convert<List<Kipon.Excel.UnitTests.Fake.Data.ValueProperty>>(mem);

                Assert.AreEqual(arrayValues.Length, result.Count);

                var ix = 0;
                foreach (var data in result)
                {
                    Assert.AreEqual(data.Boolean, arrayValues[ix].Boolean);
                    Assert.AreEqual(data.BooleanNullable, arrayValues[ix].BooleanNullable);

                    Assert.AreEqual(data.DateTime, arrayValues[ix].DateTime);
                    Assert.AreEqual(data.DateTimeNullable, arrayValues[ix].DateTimeNullable);
                    ix++;
                }
            }
        }

        public class PropertySheets
        {
            public Fake.Data.ValueProperty[] S1 { get; set; }

            public List<Fake.Data.ValueProperty> S2 { get; set; }

        }
    }
}
