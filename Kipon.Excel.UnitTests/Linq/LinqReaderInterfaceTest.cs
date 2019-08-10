using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using Kipon.Excel.Linq;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.Linq
{
    [TestClass]
    public class LinqReaderInterfaceTest
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
        public void ToArrayTest()
        {
            this.Setup();
            using (var mem = new System.IO.MemoryStream(arrayValuesData))
            {
                var result = mem.ToArray<Fake.Data.ValueProperty>();

                Assert.AreEqual(arrayValues.Length, result.Length);

            }
        }

        [TestMethod]
        public void ToObjectTest()
        {
            this.Setup();
            using (var mem = new System.IO.MemoryStream(sheetData))
            {
                var result = mem.ToObject<PropertySheets>();

                Assert.IsNotNull(result.S1);
                Assert.IsNotNull(result.S2);
            }
        }

        [TestMethod]
        public void ToListTest()
        {
            this.Setup();
            using (var mem = new System.IO.MemoryStream(arrayValuesData))
            {
                var result = mem.ToList<Fake.Data.ValueProperty>();

                Assert.AreEqual(arrayValues.Length, result.Count);

            }
        }

        public class PropertySheets
        {
            public Fake.Data.ValueProperty[] S1 { get; set; }

            public List<Fake.Data.ValueProperty> S2 { get; set; }

        }
    }
}
