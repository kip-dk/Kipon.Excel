﻿using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.WriterImplementation.Serialization
{
    internal interface IDataTypeResolver
    {
        CellValues Resolve(ICell cell);
        int? NumberOfDecimals(ICell cell);
    }
}
