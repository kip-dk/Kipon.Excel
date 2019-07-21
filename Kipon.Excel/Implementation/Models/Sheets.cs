using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Models
{
    internal class Sheets : List<Kipon.Excel.Api.ISheet>, Kipon.Excel.Implementation.Factories.IPopulator
    {
        public void Populate(object instance)
        {
            throw new NotImplementedException();
        }
    }
}
