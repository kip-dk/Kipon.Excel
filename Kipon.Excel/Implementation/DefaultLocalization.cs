﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation
{
    internal class DefaultLocalization : Api.Globalization.ILocalization
    {
        object Api.Globalization.ILocalization.ToLocal(object value)
        {
            if (value is bool)
            {
                var v = (bool)value;
                return v ? "True" : "False";
            }
            return value;
        }
    }
}
