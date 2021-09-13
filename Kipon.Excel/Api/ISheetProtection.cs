using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Api
{
    public interface ISheetProtection
    {
        SheetProtectionType Type { get; }
    }
}
