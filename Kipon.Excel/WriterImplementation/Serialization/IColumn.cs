using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.WriterImplementation.Serialization
{
    internal interface IColumn
    {
        double? Width { get; }
        int? MaxLength { get; }
        bool? Hidden { get; }
        bool? IsIndex { get; }
        object[] IndexValues { get; }
        string[] OptionSetValue { get; }
    }
}
