using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;

namespace Kipon.Excel.Implementation.Models.Sheets
{
    /// <summary>
    /// Base calss for the IEnumerable<Kipon.Excel.Api.ISheet>
    /// </summary>
    /// <typeparam name="I"></typeparam>
    internal abstract class AbstractBaseSheets<I>: IEnumerable<Kipon.Excel.Api.ISheet>, IEnumerator<ISheet>, Kipon.Excel.Implementation.Factories.IPopulator<I>
    {
        #region private fields
        private List<Kipon.Excel.Api.ISheet> _sheets = new List<ISheet>();
        private int index = -1;
        #endregion

        #region ienumerator
        public ISheet Current
        {
            get
            {
                if (index >=0 && index < _sheets.Count)
                {
                    return _sheets[index];
                }
                return null;
            }
        }

        object IEnumerator.Current
        {
            get
            {
                if (index >= 0 && index < _sheets.Count)
                {
                    return _sheets[index];
                }
                return null;
            }
        }

        public void Dispose()
        {
            this._sheets.Clear();
        }

        public bool MoveNext()
        {
            if (_sheets == null) return false;
            if (_sheets.Count == 0) return false;
            if (_sheets.Count <= (index + 1)) return false;

            index++;
            return true;
        }

        public void Reset()
        {
            this.index = -1;
        }
        #endregion

        #region abstract impl. of populate
        public abstract void Populate(I instance);
        #endregion

        #region specialized method can add sheets
        protected void Add(Kipon.Excel.Api.ISheet sheet)
        {
            if (sheet == null) throw new ArgumentNullException("sheet cannot be null");

            this._sheets.Add(sheet);
        }
        #endregion

        #region IEnumerable
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }

        public IEnumerator<ISheet> GetEnumerator()
        {
            return this;
        }
        #endregion
    }
}
