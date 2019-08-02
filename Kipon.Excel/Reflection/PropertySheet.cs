using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Reflection
{
    internal class PropertySheet
    {
        private static Dictionary<Type, PropertySheet> sheets = new Dictionary<Type, PropertySheet>();

        private List<SheetProperty> properties = new List<SheetProperty>();

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
                        sheetProperty.property = property;
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

                    if (!PropertyCell.IsCell(property.PropertyType))
                    {
                        continue;
                    }

                    var sheetProperty = new SheetProperty();
                    sheetProperty.title = property.Name;
                    sheetProperty.isReadonly = property.GetSetMethod() == null;
                    sheetProperty.property = property;
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
        internal IEnumerable<SheetProperty> Properties => this.properties;


        internal SheetProperty this[string title]
        {
            get
            {
                return (from t in this.properties where t.title == title select t).FirstOrDefault();
            }
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
            internal System.Reflection.PropertyInfo property { get; set; }
        }
    }
}
