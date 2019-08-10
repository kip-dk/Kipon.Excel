using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.Fake.Data
{
    /// <summary>
    /// S: strongly types object representing all supported value type properties than can be streamed to excel
    /// R: Added support for more value types
    /// </summary>

    public class ValueProperty
    {
        // {0} 8 chars
        private const string EMPTY_GUID_TEMLATE = "{0}-0000-0000-0000-000000000000";

        /// <summary>
        /// Needed constructor to support deserialization to this type
        /// </summary>
        public ValueProperty()
        {
        }

        /// <summary>
        /// Test constructor to create easy to validate values for all properties based on a single number
        /// </summary>
        /// <param name="data"></param>
        public ValueProperty(System.Int16 data)
        {
            this.String = $"Text {data}";

            this.Guid = new Guid(string.Format(EMPTY_GUID_TEMLATE, data.ToString().PadLeft(8,'0')));
            this.GuidNullable = this.Guid;

            this.Int16 = data;
            this.Int16Nullable = data;

            this.Int32 = data + 1000;
            this.Int32Nullable = data + 1000;

            this.Int64 = data + 100000L;
            this.Int64Nullable = data + 100000L;

            this.Decimal = data + 1.056M;
            this.DecimalNullable = data + 1.045M;

            this.Boolean = data % 2 == 0;
            this.BooleanNullable = data % 2 != 0;

            this.Enum = (EnumExample)(data % 3);
            this.EnumNullable = (EnumExample)(data % 2);

            this.DateTime = System.DateTime.Today.AddDays(data);
            this.DateTimeNullable = System.DateTime.Today.AddDays(data).AddHours(data).AddMinutes(data).AddSeconds(data);
        }

        public string String { get; set; }

        public Guid Guid { get; set; }
        public Guid? GuidNullable { get; set; }

        public System.Int16 Int16 { get; set; }
        public System.Int16? Int16Nullable { get; set; }

        public System.Int32 Int32 { get; set; }
        public System.Int32? Int32Nullable { get; set; }

        public System.Int64 Int64 { get; set; }
        public System.Int64? Int64Nullable { get; set; }

        public System.Double Double { get; set; }
        public System.Double? DoubleNullable { get; set; }

        public System.Decimal Decimal { get; set; }
        public System.Decimal? DecimalNullable { get; set; }

        public System.Boolean Boolean { get; set; }
        public System.Boolean? BooleanNullable { get; set; }

        public EnumExample Enum { get; set; }
        public EnumExample EnumNullable { get; set; }

        public DateTime DateTime { get; set; }
        public DateTime? DateTimeNullable { get; set; }


        public enum EnumExample
        {
            Value1 = 0,
            Value2 = 1,
            Value3 = 2
        }
    }
}
