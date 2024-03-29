﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;
using Kipon.Excel.WriterImplementation.Serialization;

namespace Kipon.Excel.WriterImplementation.Models.Sheet
{
    internal class PropertySheet : AbstractBaseSheet, IEnumerable<Kipon.Excel.Api.ICell>, IEnumerator<ICell>, Kipon.Excel.WriterImplementation.Serialization.IColumns
    {
        #region metadatacache
        private static Dictionary<Type, SheetMeta[]> metaCache = new Dictionary<Type, SheetMeta[]>();
        #endregion

        #region fields
        private System.Collections.IEnumerator rows;
        private int column = -1;
        private int columnPosition = -1;
        private int indexPosition = -1;
        private int row = -1;
        private SheetMeta[] sheetMetas;
        private ICell _current;
        private object _currentRow;
        private Dictionary<string, ICell> resolvedCells = new Dictionary<string, ICell>();
        private Kipon.Excel.Api.ITitleMap titleMap;
        #endregion

        #region populate impl.
        public override void Populate(object instance)
        {
            this.titleMap = null;

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

                    if (typeof(Kipon.Excel.Api.ITitleMap).IsAssignableFrom(eType))
                    {
                        this.titleMap = (Kipon.Excel.Api.ITitleMap)Activator.CreateInstance(eType);
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

                #region calculate columns for index columns
                var indexes = this.sheetMetas.Where(r => r.isIndex).ToArray();
                if (indexes.Length > 0)
                {
                    while (rows.MoveNext())
                    {
                        var row = rows.Current;
                        if (row != null)
                        {
                            foreach (var index in indexes)
                            {
                                var indexValues = index.property.GetValue(row) as IDictionary;
                                if (indexValues != null)
                                {
                                    var keys = indexValues.Keys;
                                    if (keys != null)
                                    {
                                        if (index.indexValues == null)
                                        {
                                            index.indexValues = new List<object>();
                                        }
                                        foreach (var key in keys)
                                        {
                                            if (key != null)
                                            {
                                                if (!index.indexValues.Contains(key))
                                                {
                                                    index.indexValues.Add(key);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    foreach (var index in indexes)
                    {
                        if (index.indexValues != null)
                        {
                            index._indexValues = index.indexValues.OrderBy(r => r.ToString()).ToArray();
                        }
                    }

                    this.rows.Reset();
                }
                #endregion

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
            if (this.column >= 0)
            {
                var col = this.sheetMetas[this.column];
                if (col.isIndex && this.indexPosition < (col._indexValues.Length - 1))
                {
                    this.indexPosition++;
                    this.columnPosition++;
                    this.SetCurrent();
                    return true;
                }
            }

            this.indexPosition = 0;

            if (this.column < (this.sheetMetas.Length - 1))
            {
                this.column++;
                this.columnPosition++;
                this.SetCurrent();
                return true;
            }

            this.column = 0;
            this.columnPosition = 0;

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
            var key = this.columnPosition.ToString() + ":" + (this.row + 1).ToString();
            if (this.resolvedCells.ContainsKey(key))
            {
                this._current = this.resolvedCells[key];
                return;
            }

            var meta = this.sheetMetas[this.column];
            if (this.row < 0)
            {
                var title = meta.title;

                if (meta.isIndex)
                {
                    title = meta._indexValues[this.indexPosition].ToString();
                }

                if (this.titleMap != null)
                {
                    title = this.titleMap.MapToExcelValue(title);
                }

                var cell = new Kipon.Excel.WriterImplementation.Models.Cell.Cell(this.columnPosition, this.row + 1, title);

                cell.IsHidden = meta.isHidden;
                cell.IsReadonly = true;
                cell.ValueType = typeof(string);
                cell.Decimals = meta.decimals;

                this._current = cell;
                this.resolvedCells[key] = this._current;
                return;
            }

            {
                if (meta.isIndex)
                {
                    object nextValue = null;
                    var indexColumn = meta.property.GetValue(this._currentRow) as IDictionary;

                    if (indexColumn != null)
                    {
                        var indexKey = meta._indexValues[indexPosition];
                        if (indexColumn.Contains(indexKey))
                        {
                            nextValue = indexColumn[indexKey];
                        }
                    }

                    var cell = new Kipon.Excel.WriterImplementation.Models.Cell.Cell(this.columnPosition, this.row + 1, nextValue);
                    cell.IsHidden = meta.isHidden;
                    cell.IsReadonly = meta.isReadonly;
                    cell.Decimals = meta.decimals;
                    if (nextValue != null)
                    {
                        cell.ValueType = nextValue.GetType();
                    } else
                    {
                        cell.ValueType = typeof(string);
                    }

                    this._current = cell;
                    this.resolvedCells[key] = this._current;
                    return;
                }
                else
                { 
                    var nextValue = meta.property.GetValue(this._currentRow);
                    var cell = new Kipon.Excel.WriterImplementation.Models.Cell.Cell(this.columnPosition, this.row + 1, nextValue);
                    cell.IsHidden = meta.isHidden;
                    cell.IsReadonly = meta.isReadonly;
                    cell.ValueType = meta.property.PropertyType;
                    cell.Decimals = meta.decimals;

                    if (meta.property.PropertyType.IsGenericType)
                    {
                        cell.ValueType = meta.property.PropertyType.GetGenericArguments()[0];
                    }

                    this._current = cell;
                    this.resolvedCells[key] = this._current;
                    return;
                }
            }
        }

        public void Reset()
        {
            this.column = -1;
            this.row = -1;
            this.rows.Reset();
        }
        #endregion

        #region icolumns impl
        IEnumerable<IColumn> IColumns.Columns => this.sheetMetas;
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

                if (columnAttr.MaxLength != null)
                {
                    sheetMeta.maxLength = columnAttr.MaxLength;
                }

                if (columnAttr.Width != null)
                {
                    sheetMeta.width = columnAttr.Width;
                }

                if (columnAttr.OptionSetValues != null)
                {
                    sheetMeta.optionSetValues = columnAttr.OptionSetValues;
                }

                if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Dictionary<,>))
                {
                    sheetMeta.isIndex = true;
                }
            }

            if (sheetMeta.isHidden == false)
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
                    sheetMeta.decimals = decAttr.Value;
                }
            }

            {
                var widthAttr = (Kipon.Excel.Attributes.WidthAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.WidthAttribute), true).FirstOrDefault();
                if (widthAttr != null)
                {
                    sheetMeta.width = widthAttr.Value;
                }
            }

            {
                var maxlengthAttr = (Kipon.Excel.Attributes.MaxLengthAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.MaxLengthAttribute), true).FirstOrDefault();
                if (maxlengthAttr != null)
                {
                    sheetMeta.maxLength = maxlengthAttr.Value;
                }
            }

            {
                var optionvaluesAttr = (Kipon.Excel.Attributes.OptionSetValuesAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.OptionSetValuesAttribute), true).FirstOrDefault();
                if (optionvaluesAttr != null)
                {
                    sheetMeta.optionSetValues = optionvaluesAttr.Value;
                }
            }

            {
                var titleAttr = (Kipon.Excel.Attributes.TitleAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.TitleAttribute), true).FirstOrDefault();
                if (titleAttr != null)
                {
                    sheetMeta.title = titleAttr.Value;
                }
            }

            {
                var indexAttr = (Kipon.Excel.Attributes.IndexColumnAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IndexColumnAttribute), true).FirstOrDefault();
                if (indexAttr != null)
                {
                    sheetMeta.isIndex = true;
                }
            }

            return sheetMeta;
        }
        #endregion

        #region inner classes
        private class SheetMeta : Kipon.Excel.WriterImplementation.Serialization.IColumn
        {
            internal string title { get; set; }
            internal int sort { get; set; }
            internal bool isReadonly { get; set; }
            internal bool isHidden { get; set; }
            internal int? decimals { get; set; }

            internal bool isIndex { get; set; }

            internal double? width { get; set; }
            internal int? maxLength { get; set; }
            internal string[] optionSetValues { get; set; }
            internal System.Reflection.PropertyInfo property { get; set; }

            internal List<object> indexValues { get; set; }
            internal object[] _indexValues;

            private System.Reflection.PropertyInfo keysProperty;

            internal System.Reflection.PropertyInfo KeyProperty
            {
                get
                {
                    if (keysProperty == null)
                    {
                        keysProperty = this.property.PropertyType.GetProperty("Keys");
                    }
                    return keysProperty;
                }
            }

            private System.Reflection.MethodInfo containsKey;
            internal System.Reflection.MethodInfo ContainsKey
            {
                get
                {
                    if (containsKey == null)
                    {
                        containsKey = this.property.PropertyType.GetMethod("ContainsKey");
                    }
                    return containsKey;
                }
            }

            #region icolumn impl
            double? IColumn.Width => this.width;

            int? IColumn.MaxLength => this.maxLength;

            bool? IColumn.Hidden => this.isHidden;

            bool? IColumn.IsIndex => this.isIndex;

            object[] IColumn.IndexValues => this._indexValues;

            string[] IColumn.OptionSetValue => this.optionSetValues;
            #endregion
        }
        #endregion
    }
}
