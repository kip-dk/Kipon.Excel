﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Api.Metadata
{
    internal interface IColumn
    {
        WriterImplementation.OpenXml.Types.Column Min { get; set; }
        WriterImplementation.OpenXml.Types.Column Max { get; set; }
        double Width { get; }
        bool Hidden { get; }
    }
}
