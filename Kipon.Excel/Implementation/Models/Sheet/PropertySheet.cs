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
        #region fields
        private System.Collections.IEnumerable rows;
        private int column = -1;
        private int row = -1;
        #endregion

        #region populate impl.
        public override void Populate(object instance)
        {
            if (instance is System.Collections.IEnumerable)
            {
                this.rows = (System.Collections.IEnumerable)instance;
                this.Cells = this;
                return;
            }
        }
        #endregion

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

        #region property columns initializer
        internal void AddDecorateProperties(Type type, System.Reflection.PropertyInfo[] properties)
        {
        }

        internal void AddUndecoratedProperties(Type type, System.Reflection.PropertyInfo[] properties)
        {
        }
        #endregion

        #region inner classes
        private class SheetMeta
        {
            internal string title { get; set; }
            internal int sort { get; set; }
            internal System.Reflection.PropertyInfo Property { get; set; }
        }
        #endregion
    }
}
