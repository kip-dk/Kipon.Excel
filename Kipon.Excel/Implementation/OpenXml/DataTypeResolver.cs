using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using Kipon.Excel.Api;

namespace Kipon.Excel.Implementation.OpenXml
{
    internal class DataTypeResolver : Kipon.Excel.Implementation.Serialization.IDataTypeResolver
    {
        public int? NumberOfDecimals(ICell cell)
        {
            throw new NotImplementedException();
        }

        public CellValues Resolve(ICell cell)
        {
            return CellValues.String;
        }
    }
}
