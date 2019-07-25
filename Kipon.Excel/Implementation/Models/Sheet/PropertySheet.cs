using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.Implementation.Models.Sheet
{
    internal class PropertySheet : AbstractBaseSheet, IEnumerable<Kipon.Excel.Api.ICell>, IEnumerator<ICell>
    {
        private System.Collections.IEnumerable rows;

        #region enumerator impl
        public IEnumerator<ICell> GetEnumerator()
        {
            return this;
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }
        #endregion

        public override void Populate(object instance)
        {
            if (instance is System.Collections.IEnumerable)
            {
                this.rows = (System.Collections.IEnumerable)instance;
                this.Cells = this;
                return;
            }
        }

        #region ienumerable impl
        public ICell Current => throw new NotImplementedException();

        object IEnumerator.Current => throw new NotImplementedException();

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public bool MoveNext()
        {
            throw new NotImplementedException();
        }

        public void Reset()
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}
