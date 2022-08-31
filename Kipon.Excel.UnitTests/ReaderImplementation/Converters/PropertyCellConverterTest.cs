using NUnit.Framework;
using Kipon.Excel.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.UnitTests.Fake.Data;

namespace Kipon.Excel.UnitTests.ReaderImplementation.Converters
{
    public class PropertyCellConverterTest
    {

        private Kipon.Excel.Reflection.PropertySheet ps = Kipon.Excel.Reflection.PropertySheet.ForType(typeof(Fake.Data.ValueProperty), string.Empty);

        [Test]
        public void ConverterTest()
        {
            var obj = new ValueProperty();
            var cellConverter = new Kipon.Excel.ReaderImplementation.Converters.PropertyCellConverter(obj);

            Assert.AreEqual(false, obj.Boolean);

            cellConverter.Convert(this.For(nameof(ValueProperty.Boolean)), Get("Yes"));
            Assert.AreEqual(true, obj.Boolean);

            cellConverter.Convert(this.For(nameof(ValueProperty.BooleanNullable)), Get("Yes"));
            Assert.AreEqual(true, obj.BooleanNullable);

            cellConverter.Convert(this.For(nameof(ValueProperty.Decimal)), Get("20.2"));
            Assert.AreEqual(20.2M, obj.Decimal);

            cellConverter.Convert(this.For(nameof(ValueProperty.DecimalNullable)), Get("22.2345"));
            Assert.AreEqual(22.2345M, obj.DecimalNullable);

            cellConverter.Convert(this.For(nameof(ValueProperty.Double)), Get("24.2"));
            Assert.AreEqual(24.2d, obj.Double);

            cellConverter.Convert(this.For(nameof(ValueProperty.DoubleNullable)), Get("25.2345"));
            Assert.AreEqual(25.2345d, obj.DoubleNullable);

            cellConverter.Convert(this.For(nameof(ValueProperty.Enum)), Get(ValueProperty.EnumExample.Value2.ToString()));
            Assert.AreEqual(ValueProperty.EnumExample.Value2, obj.Enum);

            cellConverter.Convert(this.For(nameof(ValueProperty.EnumNullable)), Get(ValueProperty.EnumExample.Value3.ToString()));
            Assert.AreEqual(ValueProperty.EnumExample.Value3, obj.EnumNullable);

            var g = Guid.NewGuid();
            cellConverter.Convert(this.For(nameof(ValueProperty.Guid)), Get(g.ToString()));
            Assert.AreEqual(g, obj.Guid);

            cellConverter.Convert(this.For(nameof(ValueProperty.GuidNullable)), Get(g.ToString()));
            Assert.AreEqual(g, obj.GuidNullable);

            cellConverter.Convert(this.For(nameof(ValueProperty.Int16)), Get("21"));
            Assert.AreEqual(21, obj.Int16);

            System.Int16 v = 22;
            cellConverter.Convert(this.For(nameof(ValueProperty.Int16Nullable)), Get("22"));
            Assert.AreEqual(v, obj.Int16Nullable);

            cellConverter.Convert(this.For(nameof(ValueProperty.Int32)), Get("23"));
            Assert.AreEqual(23, obj.Int32);

            cellConverter.Convert(this.For(nameof(ValueProperty.Int32Nullable)), Get("24"));
            Assert.AreEqual(24, obj.Int32Nullable);

            cellConverter.Convert(this.For(nameof(ValueProperty.Int64)), Get("25"));
            Assert.AreEqual(25L, obj.Int64);

            cellConverter.Convert(this.For(nameof(ValueProperty.Int64Nullable)), Get("26"));
            Assert.AreEqual(26L, obj.Int64Nullable);

            var d = new DateTime(2019, 08, 10, 21, 22, 24);
            var dString = d.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);

            cellConverter.Convert(this.For(nameof(ValueProperty.DateTime)), Get(dString));
            Assert.AreEqual(d, obj.DateTime);

            cellConverter.Convert(this.For(nameof(ValueProperty.DateTimeNullable)), Get(dString));
            Assert.AreEqual(d, obj.DateTimeNullable);

            cellConverter.Convert(this.For(nameof(ValueProperty.String)), Get("Some text"));
            Assert.AreEqual("Some text", obj.String);

            var now = System.DateTime.Now.AddMilliseconds(741);
            cellConverter.Convert(this.For(nameof(ValueProperty.DateTime)), new Cell { Value = now });
            Assert.AreEqual(now, obj.DateTime);

            DateTime? nowNullable = System.DateTime.Now.AddMilliseconds(743);
            cellConverter.Convert(this.For(nameof(ValueProperty.DateTime)), new Cell { Value = nowNullable });
            Assert.AreEqual(nowNullable, obj.DateTime);
        }


        private Kipon.Excel.Reflection.PropertySheet.SheetProperty For(string name)
        {
            return ps[name];
        }

        private Kipon.Excel.Api.ICell Get(string value)
        {
            return new Cell { Value = value };
        }


        public class Cell : Kipon.Excel.Api.ICell
        {
            public ICoordinate Coordinate => throw new NotImplementedException();

            public object Value { get; set; }
        }
    }
}
