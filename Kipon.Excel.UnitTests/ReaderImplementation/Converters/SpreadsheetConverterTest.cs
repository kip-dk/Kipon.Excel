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

        private PropertySheet sheet;
        private byte[] sheetData;

        private void Setup()
        {
            this.sheet = new PropertySheet();
            var numbers1 = new short[] { 0, 1, 2, 4 };

            sheet.S1 = (from number in numbers1
                        select new Fake.Data.ValueProperty(number)).ToArray();


            var numbers2 = new short[] { 7, 9, 12 };

            sheet.S2 = (from number in numbers2
                        select new Fake.Data.ValueProperty(number)).ToList();

            this.sheetData = this.sheet.ToExcel();
        }

        [TestMethod]
        public void ConvertPropertySheetTest()
        {
            this.Setup();
#warning ADD some test related to SpreadsheetConverter
        }


        public class PropertySheet
        {
            public Fake.Data.ValueProperty[] S1 { get; set; }

            public List<Fake.Data.ValueProperty> S2 { get; set; }

        }
    }
}
