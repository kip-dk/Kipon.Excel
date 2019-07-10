using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.Implementation.Serialization
{
    internal interface IDataTypeResolver
    {
        CellValues Resolve(ISheet sheet, Api.Types.Cell cell);
        int? NumberOfDecimals(ISheet sheet, Api.Types.Cell cell);
    }
}
