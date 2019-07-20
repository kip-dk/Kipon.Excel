using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.Implementation.Models
{
    public class Sheets<I>: IEnumerable<Kipon.Excel.Api.ISheet>, Kipon.Excel.Implementation.Factories.IPopulator<I>
    {
        public void Populate(I instance)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<ISheet> GetEnumerator()
        {
            throw new NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }
}
