using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Attributes
{
    public class SheetProtectionAttribute : Attribute
    {
        public SheetProtectionAttribute(Kipon.Excel.Api.SheetProtectionType type)
        {
            if (type == Api.SheetProtectionType.Custom)
            {
                throw new Kipon.Excel.Exceptions.UnsupportedTypeException("Use custom spefic constructor to define settings for custom sheet", typeof(Api.SheetProtectionType));
            }
            this.Type = type;
        }

        public SheetProtectionAttribute(string AlgorithmName, string HashValue, string SaltValue, UInt32 SpinCount)
        {
            this.AlgorithmName = AlgorithmName;
            this.HashValue = HashValue;
            this.SaltValue = SaltValue;
            this.SpinCount = SpinCount;
        }


        public Kipon.Excel.Api.SheetProtectionType Type { get; private set; }
        public string AlgorithmName { get; private set; }
        public string HashValue { get; private set; }
        public string SaltValue { get; private set; }
        public System.UInt32 SpinCount { get; private set; }
    }
}
