using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.WriterImplementation.Models.Cells
{
    internal class AbstractBaseCells : List<Kipon.Excel.Api.ICell>, Kipon.Excel.WriterImplementation.Factories.IPopulator
    {
        public void Initialize(Type type)
        {
            throw new NotImplementedException();
        }

        public void Populate(object instance)
        {
            throw new NotImplementedException();
        }
    }
}
