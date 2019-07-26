﻿using System;
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
        private System.Collections.IEnumerable rows;
        private int column = -1;
        private int row = -1;
        private SheetMeta[] sheetMetas;
        private ICell _current;
        #endregion

        #region populate impl.
        public override void Populate(object instance)
        {
            if (instance is System.Collections.IEnumerable)
            {
                var elementType = instance.GetType();
                if (elementType.IsArray)
                {
                    this.sheetMetas = metaCache[elementType.GetElementType()];
                }
                else
                {
                    if (elementType.IsGenericType)
                    {
                        this.sheetMetas = metaCache[elementType.GetGenericArguments()[0]];
                    }
                }

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
        public ICell Current => this._current;

        object IEnumerator.Current => this._current;

        public void Dispose()
        {
        }

        public bool MoveNext()
        {
            throw new NotImplementedException();
        }

        public void Reset()
        {
            this.column = -1;
            this.row = -1;
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
                isReadonly = prop.GetSetMethod() == null
            };

            var columnAttr = (Kipon.Excel.Attributes.ColumnAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.ColumnAttribute), true).FirstOrDefault();

            if (columnAttr != null)
            {
                if (!string.IsNullOrEmpty(columnAttr.Title))
                {
                    sheetMeta.title = columnAttr.Title;
                }

                if (columnAttr.Sort != null)
                {
                    sheetMeta.sort = columnAttr.Sort.Value;
                }

                if (columnAttr.IsHidden)
                {
                    sheetMeta.isHidden = true;
                }

                if (columnAttr.IsReadonly)
                {
                    sheetMeta.isReadonly = true;
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

            if (columnAttr == null || columnAttr.Sort == null)
            {
                var sortAttr = (Kipon.Excel.Attributes.SortAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.SortAttribute), true).FirstOrDefault();
                if (sortAttr != null)
                {
                    sheetMeta.sort = sortAttr.Value;
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
            internal System.Reflection.PropertyInfo property { get; set; }
        }
        #endregion
    }
}
