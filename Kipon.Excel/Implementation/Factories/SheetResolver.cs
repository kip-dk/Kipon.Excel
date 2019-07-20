using Kipon.Excel.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Factories
{
    internal class SheetResolver<T,I,J> : TypeCachedResolver<IEnumerable<ISheet>,I, Models.Sheets<I>>
        where J : T, IPopulator<I>, new()
    {
        protected override Models.Sheets<I> ResolveType(I instance)
        {
            throw new NotImplementedException();
        }
    }
}
