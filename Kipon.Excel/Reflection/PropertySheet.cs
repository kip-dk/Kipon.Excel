using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using Kipon.Excel.Extensions.Strings;
using Kipon.Excel.Exceptions;
using Kipon.Excel.Extensions;
using System.Reflection;
using DocumentFormat.OpenXml.Bibliography;

namespace Kipon.Excel.Reflection
{
    internal class PropertySheet
    {
        private static Dictionary<Type, PropertySheet> sheets = new Dictionary<Type, PropertySheet>();

        private List<SheetProperty> properties = new List<SheetProperty>();

        internal int HeaderRow { get; private set; } = 1;
        internal int HeaderColumn { get; private set; } = 1;


        internal static PropertySheet ForType(Type type)
        {
            if (sheets.ContainsKey(type))
            {
                return sheets[type];
            }
            var sheet = new PropertySheet(type);
            sheets[type] = sheet;
            return sheet;
        }

        internal static bool IsPropertySheet(Type type)
        {
            if (typeof(Kipon.Excel.Api.ISheet).IsAssignableFrom(type))
            {
                return false;
            }
            return PropertyCell.HasCell(type);
        }
        private PropertySheet(Type type)
        {
            var parser = type.GetCustomAttributes(typeof(Attributes.ParseAttribute), false).FirstOrDefault() as Attributes.ParseAttribute;
            if (parser != null)
            {
                this.HeaderRow = parser.FirstRow;
                this.HeaderColumn = parser.FirstColumn;
            }

            var properties = type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance).OrderBy(r => r.Name).ToArray();

            var ix = int.MaxValue - (properties.Length + 1);

            #region column decorated properties
            {
                foreach (var property in properties)
                {
                    var ignore = (Kipon.Excel.Attributes.IgnoreAttribute)property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IgnoreAttribute), true).FirstOrDefault();
                    if (ignore != null)
                    {
                        continue;
                    }

                    var columnAttr = (Kipon.Excel.Attributes.ColumnAttribute)property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.ColumnAttribute), true).FirstOrDefault();
                    if (columnAttr != null)
                    {
                        var sheetProperty = new SheetProperty();
                        sheetProperty.title = columnAttr.Title ?? property.Name;
                        sheetProperty.decimals = columnAttr.Decimals;
                        sheetProperty.isHidden = columnAttr.IsHidden;
                        sheetProperty.isReadonly = columnAttr.IsReadonly;
                        sheetProperty.maxlength = columnAttr.MaxLength;
                        sheetProperty.optionSetValues = columnAttr.OptionSetValues;
                        sheetProperty.sort = columnAttr.Sort > int.MinValue ? columnAttr.Sort : ix++;
                        sheetProperty.width = columnAttr.Width;
                        sheetProperty.maxlength = columnAttr.MaxLength;
                        sheetProperty.optional = columnAttr.Optional;
                        sheetProperty.property = property;

                        sheetProperty.ResolveIndexer(type);
                        sheetProperty.ResolveAlias();

                        this.Populate(sheetProperty);
                        this.properties.Add(sheetProperty);
                    }
                }

                if (this.properties.Count > 0)
                {
                    this.properties = this.properties.OrderBy(r => r.sort).ToList();
                    return;
                }
            }
            #endregion

            #region duck type properties
            {
                foreach (var property in properties)
                {
                    var ignore = (Kipon.Excel.Attributes.IgnoreAttribute)property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IgnoreAttribute), true).FirstOrDefault();
                    if (ignore != null)
                    {
                        continue;
                    }

                    var isIndex = property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IndexColumnAttribute), false).Any();

                    if (!isIndex && !PropertyCell.IsCell(property.PropertyType))
                    {
                        continue;
                    }

                    var sheetProperty = new SheetProperty();
                    sheetProperty.title = property.Name;
                    sheetProperty.isReadonly = property.GetSetMethod() == null;
                    sheetProperty.property =  property;
                    sheetProperty.optional = property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.OptionalAttribute), false).Any();
                    sheetProperty.ResolveIndexer(type);
                    sheetProperty.ResolveAlias();

                    this.Populate(sheetProperty);
                    this.properties.Add(sheetProperty);
                }

                if (this.properties.Count > 0)
                {
                    this.properties = this.properties.OrderBy(r => r.sort).ToList();
                    return;
                }
            }
            #endregion
        }


        private void Populate(SheetProperty property)
        {
            if (property.isHidden == false)
            {
                var hiddenAttr = (Kipon.Excel.Attributes.HiddenAttribute)property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.HiddenAttribute), true).FirstOrDefault();
                if (hiddenAttr != null)
                {
                    property.isHidden = true;
                }
            }

            if (property.isReadonly == false)
            {
                var readonlyAttr = (Kipon.Excel.Attributes.ReadonlyAttribute)property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.ReadonlyAttribute), true).FirstOrDefault();
                if (readonlyAttr != null)
                {
                    property.isReadonly = true;
                }
            }

            {
                var sortAttr = (Kipon.Excel.Attributes.SortAttribute)property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.SortAttribute), true).FirstOrDefault();
                if (sortAttr != null)
                {
                    property.sort = sortAttr.Value;
                }
            }

            {
                var decAttr = (Kipon.Excel.Attributes.DecimalsAttribute)property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.DecimalsAttribute), true).FirstOrDefault();
                if (decAttr != null)
                {
                    property.decimals = decAttr.Value;
                }
            }

            {
                var widthAttr = (Kipon.Excel.Attributes.WidthAttribute)property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.WidthAttribute), true).FirstOrDefault();
                if (widthAttr != null)
                {
                    property.width = widthAttr.Value;
                }
            }

            {
                var maxlengthAttr = (Kipon.Excel.Attributes.MaxLengthAttribute)property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.MaxLengthAttribute), true).FirstOrDefault();
                if (maxlengthAttr != null)
                {
                    property.maxlength = maxlengthAttr.Value;
                }
            }

            {
                var optionvaluesAttr = (Kipon.Excel.Attributes.OptionSetValuesAttribute)property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.OptionSetValuesAttribute), true).FirstOrDefault();
                if (optionvaluesAttr != null)
                {
                    property.optionSetValues = optionvaluesAttr.Value;
                }
            }

            {
                var titleAttr = (Kipon.Excel.Attributes.TitleAttribute)property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.TitleAttribute), true).FirstOrDefault();
                if (titleAttr != null)
                {
                    property.title = titleAttr.Value;
                }
            }
        }

        private Dictionary<string, SheetProperty> indexMap = new Dictionary<string, SheetProperty>();

        internal SheetProperty GetIndexMatch(string name, int? columnIndex = null)
        {
            if (indexMap.ContainsKey(name))
            {
                var map = indexMap[name];
                if (columnIndex != null)
                {
                    map[columnIndex.Value] = name;
                }
                return map;
            }

            var indexers = (from p in this.properties
                            where p.indexExpression != null
                            select p).ToArray();

            foreach (var index in indexers)
            {
                var match = index.indexExpression.Match(name);
                if (match.Success && match.Value == name && match.NextMatch().Success == false)
                {
                    indexMap.Add(name, index);

                    if (columnIndex.HasValue)
                    {
                        index[columnIndex.Value] = name;
                    }
                    return index;
                }
            }

            return null;
        }

        internal IEnumerable<SheetProperty> Properties => this.properties;


        internal SheetProperty this[string title]
        {
            get
            {
                return (from t in this.properties where t.title == title select t).FirstOrDefault();
            }
        }

        internal Kipon.Excel.Api.ISheet BestMatch(IEnumerable<Kipon.Excel.Api.ISheet> sheets)
        {
            var firstRowNumber = this.HeaderRow - 1;
            var firstColumnNumber = this.HeaderColumn - 1;

            foreach (var sheet in sheets)
            {
                var firstRow = (from c in sheet.Cells
                                where c.Coordinate.Point.Last() == firstRowNumber
                                  && c.Coordinate.Point.First() >= firstColumnNumber
                                  && c.Value != null
                                select c.Value.ToString()).ToArray();

                if (firstRow != null && firstRow.Length > 0)
                {
                    var match = 0;
                    foreach (var f in firstRow)
                    {
                        var prop = this.Match(f);
                        if (prop != null)
                        {
                            match++;
                            continue;
                        }
                    }

                    if (firstRow.Length == match)
                    {
                        // 100% match
                        return sheet;
                    }

                    var foundfields = firstRow.Select(r => r.ToRelaxedName()).ToArray();
                    var missings = this.properties.Where(r => !r.optional && !foundfields.Contains(r.title.ToRelaxedName())).Any();

                    if (!missings)
                    {
                        return sheet;
                    }
                }
            }
            return null;
        }

        internal SheetProperty Match(string name, int? columnIndexNumber = null)
        {
            var prop = this.properties.Where(p => p.title == name).SingleOrDefault();
            if (prop != null)
            {
                return prop;
            }

            prop = this.properties.Where(p => p.title.ToRelaxedName() == name.ToRelaxedName()).SingleOrDefault();
            if (prop != null)
            {
                return prop;
            }

            prop = this.properties.Where(p => p.Alias != null && p.Alias.Contains(name)).SingleOrDefault();
            if (prop != null)
            {
                return prop;
            }

            prop = this.properties.Where(p => p.Alias != null && p.Alias.Select(r => r.ToRelaxedName()).Contains(name.ToRelaxedName())).SingleOrDefault();
            if (prop != null)
            {
                return prop;
            }

            prop = this.GetIndexMatch(name, columnIndexNumber);
            if (prop != null)
            {
                return prop;
            }

            return null;
        }

        internal class SheetProperty
        {
            internal string title { get; set; }
            internal int sort { get; set; }
            internal bool isHidden { get; set; }
            internal bool isReadonly { get; set; }
            internal int? decimals { get; set; }
            internal int? maxlength { get; set; }
            internal string[] optionSetValues { get; set; }
            internal double? width { get; set; }
            internal bool optional { get; set; }

            internal string indexPattern { get; set; }
            internal Type indexType { get; set; }
            internal MethodInfo indexAddMethod { get; set; }
            internal System.Text.RegularExpressions.Regex indexExpression;
            internal string[] Alias { get; set; }

            private System.Reflection.PropertyInfo _property;
            internal System.Reflection.PropertyInfo property
            {
                get
                {
                    return this._property;
                }
                set
                {
                    this._property = value;
                    this.canWrite = this._property.GetSetMethod() != null;
                    this.PropertyType = Nullable.GetUnderlyingType(this._property.PropertyType) ?? this._property.PropertyType;
                }
            }

            private Dictionary<int, string> indexTitleMap = new Dictionary<int, string>();

            public string this[int index]
            {
                get
                {
                    return this.indexTitleMap[index];
                }
                set
                {
                    this.indexTitleMap[index] = value;
                }
            }

            internal bool canWrite { get; private set; }

            // returns the flattner type of the property, so if is Nullable<int>, int will be returned
            internal Type PropertyType { get; private set; }

            internal void ResolveAlias()
            {
                var aliasAttributes = this.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.AliasAttribute), false).ToArray();
                if (aliasAttributes.Length > 0)
                {
                    this.Alias = new string[aliasAttributes.Count()];
                    var ix = 0;
                    foreach (Kipon.Excel.Attributes.AliasAttribute a in aliasAttributes)
                    {
                        this.Alias[ix] = a.Title;
                    }
                }
            }

            internal void ResolveIndexer(Type sheetType)
            {
                var indexProperty = this.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IndexColumnAttribute), false).FirstOrDefault() as Kipon.Excel.Attributes.IndexColumnAttribute;
                if (indexProperty != null)
                {
                    bool isDict = this.property.PropertyType.IsGenericType && this.property.PropertyType.GetGenericTypeDefinition() == typeof(Dictionary<,>);

                    if (!isDict)
                    {
                        throw new Exceptions.UnsupportedColumnIndexPropertyException(sheetType, this.property);
                    }

                    this.indexPattern = indexProperty.Pattern;
                    this.indexExpression = indexProperty.Expression;

                    var args = this.PropertyType.GetGenericArguments();

                    if (args[0] != typeof(string))
                    {
                        throw new Exceptions.UnsupportedColumnIndexPropertyException(sheetType, this.property);
                    }
                    this.indexType = args[1];

                    if (!PropertyCell.IsCell(this.indexType))
                    {
                        throw new Exceptions.UnsupportedColumnIndexPropertyException(sheetType, this.property);
                    }

                    this.indexAddMethod = this.property.PropertyType.GetMethod("Add", new Type[] { typeof(string), this.indexType });
                }
            }

            internal object ValueOf(object value)
            {
                if (this.indexExpression != null)
                {
                    return value.ToValue(this.indexType);
                }
                return value.ToValue(this.PropertyType);
            }
        }
    }
}
