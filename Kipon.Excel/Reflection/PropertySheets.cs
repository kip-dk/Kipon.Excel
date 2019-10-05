using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.Reflection
{
    internal class PropertySheets
    {
        private static Dictionary<Type, PropertySheets> resolved = new Dictionary<Type, PropertySheets>();
        private List<SheetsProperty> properties = new List<SheetsProperty>();

        internal static PropertySheets ForType(Type type)
        {
            if (resolved.ContainsKey(type))
            {
                return resolved[type];
            }

            if (!IsPropertySheets(type))
            {
                return null;
            }

            var result = new PropertySheets(type);
            resolved.Add(type, result);

            return result;
        }

        internal static bool IsPropertySheets(Type type)
        {
            if (typeof(Kipon.Excel.Api.ISpreadsheet).IsAssignableFrom(type))
            {
                return false;
            }

            var properties = type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            foreach (var prop in properties)
            {
                if (typeof(Kipon.Excel.Api.ISheet).IsAssignableFrom(prop.PropertyType))
                {
                    return true;
                }

                if (PropertySheet.IsPropertySheet(prop.PropertyType))
                {
                    return true;
                }
            }

            return false;
        }

        internal IEnumerable<SheetsProperty> Properties => this.properties;

        public SheetsProperty this[string title]
        {
            get
            {
                if (string.IsNullOrEmpty(title))
                {
                    return null;
                }
                return (from p in this.properties where p.Title == title select p).FirstOrDefault();
            }
        }

        private PropertySheets(Type type)
        {
            var properties = type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance).OrderBy(r => r.Name).ToArray();

            var ix = int.MaxValue - (properties.Length + 1);
            foreach (var property in properties)
            {
                var ignore = (Kipon.Excel.Attributes.IgnoreAttribute)property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IgnoreAttribute), true).FirstOrDefault();
                if (ignore != null)
                {
                    continue;
                }

                var sheetAttr = (Kipon.Excel.Attributes.SheetAttribute)property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.SheetAttribute), true).FirstOrDefault();
                if (sheetAttr != null)
                {
                    var next = new SheetsProperty();
                    next.Title = sheetAttr.Title ?? property.Name;
                    next.Sort = sheetAttr.Sort ?? ix++;
                    next.property = property;
                    next.ElementType = property.PropertyType.GetElementType();
                    next.IsReadyonly = property.GetSetMethod() == null;
                    this.properties.Add(next);
                    this.Populate(next);
                    continue;
                }
            }

            if (this.properties.Count > 0)
            {
                this.properties = this.properties.OrderBy(r => r.Sort).ToList();
                return;
            }

            foreach (var property in properties)
            {
                var ignore = (Kipon.Excel.Attributes.IgnoreAttribute)property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.IgnoreAttribute), true).FirstOrDefault();
                if (ignore != null)
                {
                    continue;
                }

                if (property.PropertyType.IsArray && PropertySheet.IsPropertySheet(property.PropertyType.GetElementType()))
                {
                    var next = new SheetsProperty();
                    next.Title = property.Name;
                    next.property = property;
                    next.ElementType = property.PropertyType.GetElementType();
                    next.IsReadyonly = property.GetSetMethod() == null;
                    this.properties.Add(next);
                    this.Populate(next);
                    continue;
                }

                if (property.PropertyType.IsGenericType && typeof(System.Collections.IEnumerable).IsAssignableFrom(property.PropertyType) && PropertySheet.IsPropertySheet(property.PropertyType.GetGenericArguments()[0]))
                {
                    var next = new SheetsProperty();
                    next.Title = property.Name;
                    next.property = property;
                    next.ElementType = property.PropertyType.GetGenericArguments()[0];
                    this.properties.Add(next);
                    this.Populate(next);
                    continue;
                }
            }

            if (this.properties.Count > 0)
            {
                this.properties = this.properties.OrderBy(r => r.Sort).ToList();
                return;
            }
        }

        private void Populate(SheetsProperty property)
        {
            {
                var sortAttr = (Kipon.Excel.Attributes.SortAttribute)property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.SortAttribute), true).FirstOrDefault();
                if (sortAttr != null)
                {
                    property.Sort = sortAttr.Value;
                }
            }
            {
                var titleAttr = (Kipon.Excel.Attributes.TitleAttribute)property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.TitleAttribute), true).FirstOrDefault();
                if (titleAttr != null)
                {
                    property.Title = titleAttr.Value;
                }
            }
            {
                if (property.IsReadyonly == false)
                {
                    property.IsReadyonly = property.property.GetCustomAttributes(typeof(Kipon.Excel.Attributes.ReadonlyAttribute), false).Any();

                }
            }
        }

        internal class SheetsProperty
        {
            internal string Title { get; set; }
            internal int Sort { get; set; }
            internal bool IsReadyonly { get; set; }
            internal System.Reflection.PropertyInfo property { get; set; }
            internal Type ElementType { get; set; }

            internal Kipon.Excel.Api.ISheet BestMatch(IEnumerable<Kipon.Excel.Api.ISheet> sheets)
            {
                if (sheets == null)
                {
                    return null;
                }

                var titleMatch = (from t in sheets where t.Title == this.Title select t).ToArray();
                if (titleMatch.Length == 1)
                {
                    return titleMatch[0];
                }

                return null;
            }
        }
    }
}
