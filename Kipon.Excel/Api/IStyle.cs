using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Api
{
    public interface IStyle
    {
        System.UInt32? Resolve(Kipon.Excel.Api.ICell cell);
    }
}
