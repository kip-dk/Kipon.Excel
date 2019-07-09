using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Serialization
{
    internal interface IDataTypeResolver
    {
        CellValues Resolve(ISheet sheet, Types.Cell cell);
        int? NumberOfDecimals(ISheet sheet, Types.Cell cell);
    }
}
