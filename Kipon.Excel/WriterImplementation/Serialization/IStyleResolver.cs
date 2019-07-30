using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.WriterImplementation.Serialization
{
    internal interface IStyleResolver
    {
        uint Resolve(Kipon.Excel.Api.ICell cell);
    }
}
