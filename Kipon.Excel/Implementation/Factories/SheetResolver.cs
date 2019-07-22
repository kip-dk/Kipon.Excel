using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Implementation.Models.Sheet;

namespace Kipon.Excel.Implementation.Factories
{
    internal class SheetResolver : TypeCachedResolver<Kipon.Excel.Api.ISheet, Models.Sheet.AbstractBaseSheet>
    {
        protected override AbstractBaseSheet ResolveType(Type instanceType)
        {
            throw new NotImplementedException();
        }
    }
}
