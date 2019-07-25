using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Implementation.Models.Sheets
{
    internal class SingleSheets : AbstractBaseSheets
    {
        public override void Populate(object instance)
        {
            var sheetResolver = new Kipon.Excel.Implementation.Factories.SheetResolver();

            var sheet = sheetResolver.Resolve(instance);
        }
    }
}
