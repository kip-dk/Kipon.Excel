using Kipon.Excel.Attributes;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;
using Kipon.Excel.Linq;

namespace Kipon.Excel.UnitTests.TempTests
{
    public class EnvelopeThing
    {
        // private const string excelfilename = @"C:\Projects\Tagion\DocuSign\Envelopes-beriget.xlsx";


        [Test]
        public void EnvelopeTest()
        {
            /*
            using (var fs = new System.IO.FileStream(excelfilename, System.IO.FileMode.Open))
            {
                var envelopes =  fs.ToArray<ResultEnvelope>();


                var count = envelopes.Length;
            }
            */
        }


        public class ResultEnvelope
        {
            [Sort(1)]
            [Hidden]
            public string AccountId { get; set; }

            [Sort(2)]
            public string EnvelopeId { get; set; }

            [Sort(3)]
            public string AccountName { get; set; }

            [Sort(4)]
            public string Email { get; set; }

            [Sort(5)]
            public string Email2 { get; set; }

            [Sort(6)]
            public string Email3 { get; set; }

            [Sort(7)]
            public string Email4 { get; set; }

            [Sort(8)]
            public string Name { get; set; }

            [Sort(9)]
            public decimal? Tagions { get; set; }

            [Sort(10)]
            public decimal? Amount { get; set; }

            [Sort(11)]
            public DateTime SendDate { get; set; }

            [Sort(12)]
            public DateTime SignedDate { get; set; }

            [Sort(13)]
            public string Address { get; set; }

            [Sort(14)]
            public string Country { get; set; }

            [Sort(15)]
            [Optional]
            public string ImportStatus { get; set; }

            [Sort(16)]
            [Optional]
            public string ImportMessage { get; set; }

            [Sort(17)]
            [Optional]
            public int? CrmId { get; set; }

        }

    }
}
