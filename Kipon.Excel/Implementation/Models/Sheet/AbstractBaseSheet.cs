using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.Implementation.Models.Sheet
{
    internal abstract class AbstractBaseSheet : Kipon.Excel.Api.ISheet, Kipon.Excel.Implementation.Factories.IPopulator
    {
        public string Title { get; set; }

        public IEnumerable<ICell> Cells { get; protected set; }

        public abstract void Populate(object instance);
    }
}
