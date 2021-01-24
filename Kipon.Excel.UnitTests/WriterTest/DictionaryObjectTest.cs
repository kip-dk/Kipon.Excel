using Kipon.Excel.Attributes;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;
using Kipon.Excel.Linq;

namespace Kipon.Excel.UnitTests.WriterTest
{
    public class DictionaryObjectTest
    {
        [Test]
        public void TestObjectWithDictionary()
        {
            var rows = new DicClass[3];

            rows[0] = new DicClass
            {
                Name = "R1",
                ["A-Date"] = System.DateTime.Now,
                ["A-int"] = 10,
                ["A-decimal"] = 12.2M
            };

            rows[1] = new DicClass
            {
                Name = "R2",
                ["A-Date"] = System.DateTime.Now.AddHours(10),
                ["A-int"] = 20,
                ["A-decimal"] = 22.2M
            };

            rows[2] = new DicClass
            {
                Name = "R3"
            };

            var excel = rows.ToExcel();

            System.IO.File.WriteAllBytes(@"C:\Temp\dic.xlsx", excel);

        }


        public class DicClass
        {
            public DicClass()
            {
            }

            [Sort(1)]
            [Title("Name")]
            public string Name { get; set; }

            [Sort(2)]
            [IndexColumn("[AB]-.*")]
            public Dictionary<string, object> Fields { get; set; }

            internal object this[string name]
            {
                get
                {
                    if (this.Fields.TryGetValue(name, out object v)) return v;
                    return null;
                }
                set
                {
                    if (this.Fields == null)
                    {
                        this.Fields = new Dictionary<string, object>();
                    }
                    this.Fields[name] = value;
                }
            }
        }
    }
}
