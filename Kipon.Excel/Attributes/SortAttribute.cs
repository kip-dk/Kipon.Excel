﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class SortAttribute : Attribute
    {
        public SortAttribute(int value)
        {
            this.Value = value;
        }

        public int Value { get; private set; }
    }
}
