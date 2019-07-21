using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Models
{
    internal class Sheets<I> : List<Kipon.Excel.Api.ISheet>, Kipon.Excel.Implementation.Factories.IPopulator<I>
    {
        public void Populate(I instance)
        {
            throw new NotImplementedException();
        }
    }
}
