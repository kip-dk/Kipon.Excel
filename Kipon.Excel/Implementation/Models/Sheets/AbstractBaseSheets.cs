using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Models.Sheets
{
    internal abstract class AbstractBaseSheets : List<Kipon.Excel.Api.ISheet>, Kipon.Excel.Implementation.Factories.IPopulator
    {
        public abstract void Populate(object instance);
    }
}
