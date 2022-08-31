using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.ReaderImplementation.Converters
{
    public class PropertySheetConverterTest
    {
        [Test]
        public void ConvertTest()
        {
            var numbers = new short[] { 1, 2, 3, 4, 5 };
            var rows = (from number in numbers
                        select new Fake.Data.ValueProperty(number)).ToArray();

            var spreadsheetResolver = new Kipon.Excel.WriterImplementation.Factories.SpreadsheetResolver();
            var spreadsheet = spreadsheetResolver.Resolve(rows);

            var ps = Kipon.Excel.Reflection.PropertySheet.ForType(typeof(Fake.Data.ValueProperty), string.Empty);
            var psConverter = new Kipon.Excel.ReaderImplementation.Converters.PropertySheetConverter(typeof(Fake.Data.ValueProperty), ps);

            var result = psConverter.Convert(spreadsheet.Sheets.First());

            Assert.AreEqual(5, result.Count);
            var returnType = typeof(List<Fake.Data.ValueProperty>);

            Assert.AreEqual(returnType, result.GetType());

            Assert.AreNotEqual(rows.First(), result[0]);


            var ix = 0;
            foreach (var input in rows)
            {
                var output = (Fake.Data.ValueProperty)result[ix];

                Assert.AreEqual(input.Boolean, output.Boolean);
                Assert.AreEqual(input.BooleanNullable, output.BooleanNullable);

                Assert.AreEqual(input.DateTime, output.DateTime);
                Assert.AreEqual(input.DateTimeNullable, output.DateTimeNullable);

                Assert.AreEqual(input.Decimal, output.Decimal);
                Assert.AreEqual(input.DecimalNullable, output.DecimalNullable);

                Assert.AreEqual(input.Double, output.Double);
                Assert.AreEqual(input.DoubleNullable, output.DoubleNullable);

                Assert.AreEqual(input.Enum, output.Enum);
                Assert.AreEqual(input.EnumNullable, output.EnumNullable);

                Assert.AreEqual(input.Guid, output.Guid);
                Assert.AreEqual(input.GuidNullable, output.GuidNullable);

                Assert.AreEqual(input.Int16, output.Int16);
                Assert.AreEqual(input.Int16Nullable, output.Int16Nullable);

                Assert.AreEqual(input.Int32, output.Int32);
                Assert.AreEqual(input.Int32Nullable, output.Int32Nullable);

                Assert.AreEqual(input.Int64, output.Int64);
                Assert.AreEqual(input.Int64Nullable, output.Int64Nullable);

                Assert.AreEqual(input.String, output.String);
                ix++;
            }
        }
    }
}
