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
        public ValueProperty(int data)
        {
            this.Guid = new Guid(string.Format(EMPTY_GUID_TEMLATE, data.ToString().PadLeft(8)));
            this.GuidNullable = this.Guid;

            this.Int = data;
            this.IntNullable = data;
        }

        public Guid Guid { get; set; }
        public Guid? GuidNullable { get; set; }

        public int Int { get; set; }
        public int? IntNullable { get; set; }

#warning not all types has been added
    }
}
