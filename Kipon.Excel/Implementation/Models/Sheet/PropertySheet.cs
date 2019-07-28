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
        #region metadatacache
        private static Dictionary<Type, SheetMeta[]> metaCache = new Dictionary<Type, SheetMeta[]>();
        #endregion

        #region fields
        private System.Collections.IEnumerator rows;
        private int column = -1;
        private int row = -1;
        private SheetMeta[] sheetMetas;
        private ICell _current;
        private object _currentRow;
        private Dictionary<string, ICell> resolvedCells = new Dictionary<string, ICell>();
        #endregion

        #region populate impl.
        public override void Populate(object instance)
        {
            if (instance is System.Collections.IEnumerable)
            {
                var elementType = instance.GetType();
                if (elementType.IsArray)
                {
                    var eType = elementType.GetElementType();
                    this.sheetMetas = metaCache[eType];
                    if (string.IsNullOrEmpty(this.Title))
                    {
                        this.Title = eType.Name;
                    }
                }
                else
                {
                    if (elementType.IsGenericType)
                    {
                        var eType = elementType.GetGenericArguments()[0];
                        this.sheetMetas = metaCache[eType];
                        if (string.IsNullOrEmpty(this.Title))
                        {
                            this.Title = eType.Name;
                        }
                    }
                }

                this.rows = ((System.Collections.IEnumerable)instance).GetEnumerator();
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
        public ICell Current => this._current;

        object IEnumerator.Current => this._current;

        public void Dispose()
        {
        }

        public bool MoveNext()
        {
            if (this.column < (this.sheetMetas.Length - 1))
            {
                this.column++;
                this.SetCurrent();
                return true;
            }

            this.column = 0;
            var hasNext = this.rows.MoveNext();
            if (hasNext)
            {
                row++;
                this._currentRow = this.rows.Current;
                this.SetCurrent();
                return true;
            }
            return false;
        }

        private void SetCurrent()
        {
            var key = this.column.ToString() + ":" + (this.row + 1).ToString();
            if (this.resolvedCells.ContainsKey(key))
            {
                this._current = this.resolvedCells[key];
                return;
            }

            var meta = this.sheetMetas[this.column];
            if (this.row < 0)
            {
                var cell = new Kipon.Excel.Implementation.Models.Cell.Cell(this.column, this.row + 1, meta.title);
                cell.IsHidden = meta.isHidden;
                cell.IsReadonly = true;
                cell.ValueType = typeof(string);

                this._current = cell;
                this.resolvedCells[key] = this._current;
                return;
            }

            {
                var nextValue = meta.property.GetValue(this._currentRow);
                var cell = new Kipon.Excel.Implementation.Models.Cell.Cell(this.column, this.row + 1, nextValue);
                cell.IsHidden = meta.isHidden;
                cell.IsReadonly = meta.isReadonly;
                cell.ValueType = meta.property.PropertyType;

                this._current = cell;
                this.resolvedCells[key] = this._current;
                return;
            }
        }

        public void Reset()
        {
            this.column = -1;
            this.row = -1;
            this.rows.Reset();
        }
        #endregion

        #region property columns initializer
        internal void AddPropertyInfos(Type type, System.Reflection.PropertyInfo[] properties)
        {
            var sheetMetas = new List<SheetMeta>();
            var ix = int.MaxValue - (properties.Length - 1);
            foreach (var prop in properties.OrderBy(r => r.Name))
            {
                var next = CreateSheetMeta(prop, ix);
                sheetMetas.Add(next);
                ix++;
            }
            metaCache[type] = sheetMetas.OrderBy(r => r.sort).ToArray();
        }
        #endregion

        #region helpermethod
        private SheetMeta CreateSheetMeta(System.Reflection.PropertyInfo prop, int ix)
        {
            var sheetMeta = new SheetMeta
            {
                title = prop.Name,
                sort = ix,
                isHidden = false,
                isReadonly = prop.GetSetMethod() == null,
                property = prop
            };

            var columnAttr = (Kipon.Excel.Attributes.ColumnAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.ColumnAttribute), true).FirstOrDefault();

            if (columnAttr != null)
            {
                if (!string.IsNullOrEmpty(columnAttr.Title))
                {
                    sheetMeta.title = columnAttr.Title;
                }

                if (columnAttr.Sort != int.MinValue)
                {
                    sheetMeta.sort = columnAttr.Sort;
                }

                if (columnAttr.IsHidden)
                {
                    sheetMeta.isHidden = true;
                }

                if (columnAttr.IsReadonly)
                {
                    sheetMeta.isReadonly = true;
                }

                if (columnAttr.Decimals != null)
                {
                    sheetMeta.decimals = columnAttr.Decimals.Value;
                }
            }

            if (sheetMeta.isReadonly == false)
            {
                var hiddenAttr = (Kipon.Excel.Attributes.HiddenAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.HiddenAttribute), true).FirstOrDefault();
                if (hiddenAttr != null)
                {
                    sheetMeta.isHidden = true;
                }
            }

            if (sheetMeta.isReadonly == false)
            {
                var readonlyAttr = (Kipon.Excel.Attributes.ReadonlyAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.ReadonlyAttribute), true).FirstOrDefault();
                if (readonlyAttr != null)
                {
                    sheetMeta.isReadonly = true;
                }
            }

            {
                var sortAttr = (Kipon.Excel.Attributes.SortAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.SortAttribute), true).FirstOrDefault();
                if (sortAttr != null)
                {
                    sheetMeta.sort = sortAttr.Value;
                }
            }

            {
                var decAttr = (Kipon.Excel.Attributes.DecimalsAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.DecimalsAttribute), true).FirstOrDefault();
                if (decAttr != null)
                {
                    sheetMeta.sort = decAttr.Value;
                }
            }
            return sheetMeta;
        }
        #endregion

        #region inner classes
        private class SheetMeta
        {
            internal string title { get; set; }
            internal int sort { get; set; }
            internal bool isReadonly { get; set; }
            internal bool isHidden { get; set; }
            internal int? decimals { get; set; }
            internal System.Reflection.PropertyInfo property { get; set; }
        }
        #endregion
    }
}
