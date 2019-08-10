using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;


namespace Kipon.Excel.UnitTests.Linq
{
    [TestClass]
    public class LinqWriterInterfaceTest
    {
        [TestMethod]
        public void ToExcelKnownArrayTypeTest()
        {
            var data = new System.Int16[] { 1, 2, 3 };
            
            var exceldata = (from d in data
                            select new Kipon.Excel.UnitTests.Fake.Data.ValueProperty(d)
                            ).ToArray();

            var excel = exceldata.ToExcel();

            var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
            using (var mem = new System.IO.MemoryStream(excel))
            {
                using (var excelDoc = SpreadsheetDocument.Open(mem, false))
                {
                    var errors = validator.Validate(excelDoc).Where(r => !r.Description.StartsWith("The attribute 'rgb' has invalid value")).ToArray();
                    Assert.AreEqual(0, errors.Length);
                }
            }

        }


        [TestMethod]
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
                             }).ToArray().ToExcel();

            var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
            using (var mem = new System.IO.MemoryStream(exceldata))
            {
                using (var excelDoc = SpreadsheetDocument.Open(mem, false))
                {
                    var errors = validator.Validate(excelDoc).Where(r => !r.Description.StartsWith("The attribute 'rgb' has invalid value")).ToArray();
                    Assert.AreEqual(0, errors.Length);
                }
            }
        }
    }
}
