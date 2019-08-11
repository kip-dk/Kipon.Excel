﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Api
{
    public interface ICell
    {
        ICoordinate Coordinate { get; }
        object Value { get; }
    }
}
