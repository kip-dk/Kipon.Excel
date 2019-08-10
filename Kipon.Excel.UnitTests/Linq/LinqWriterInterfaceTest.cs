using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Linq;

namespace Kipon.Excel.UnitTests.Linq
{
    [TestClass]
    public class LinqWriterInterfaceTest
    {
        //[TestMethod]
        public void ToExcelKnownArrayTypeTest()
        {
            var data = new System.Int16[] { 1, 2, 3 };
            
            var exceldata = from d in data
                            select new Kipon.Excel.UnitTests.Fake.Data.ValueProperty(d);

            var excel = exceldata.ToExcel();
        }


        //[TestMethod]
        public void ToExcelAnnonumousArrayTest()
        {
            var  data = new int[] { 1, 2, 3 };

            var exceldata = (from d in data
                             select new
                             {
                                 Id = d,
                                 Name = $"Name of {d}",
                                 IsSomthing = d % 2 == 0,
                                 Date = new System.DateTime(2019,11, d),
                                 DecimalNumber = (decimal)d + 12.56M
                             }).ToExcel();

            Assert.IsNotNull(exceldata);
            Assert.IsTrue(exceldata.Length > 0);
        }
    }
}
