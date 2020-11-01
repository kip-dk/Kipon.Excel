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

namespace Kipon.Excel.UnitTests.Api
{
    public class ExcelStreamTest
    {
        [Test]
        public void LogTest()
        {
            using (var str = this.GetType().Assembly.GetManifestResourceStream("Kipon.Excel.UnitTests.Resources.LogTest.xlsx"))
            {
                using (var excelStream = new Kipon.Excel.Api.ExcelStream(str))
                {
                    var log = new Log();
                    excelStream.Log = log;
                    var result = excelStream.ToArray<Data>();
                    Assert.AreEqual(1, log.logs.Count);
                }
            }
        }


        public class Data
        {
            public string C1 { get; set; }
            public string C2 { get; set; }
        }


        internal class Log : Kipon.Excel.Api.ILog
        {
            internal List<string> logs = new List<string>();

            void ILog.Log(string message)
            {
                logs.Add(message);
            }
        }
    }
}
