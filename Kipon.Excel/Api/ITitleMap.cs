using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Api
{
    public interface ITitleMap
    {
        string MapToExcelValue(string value);
        string MapFromExcelValue(string value);
    }
}
