using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.Api
{
    /// <summary>
    /// Let your Sheet class implement this interface if you need a reference to what line number in excel these data is representing.
    /// That can be convinient when reporting errors or similar to the user.
    /// </summary>
    public interface IRowAware
    {
        void SetRowNo(int excelRowNumber);
    }
}
