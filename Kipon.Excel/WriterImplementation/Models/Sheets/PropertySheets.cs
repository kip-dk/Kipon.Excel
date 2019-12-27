using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.WriterImplementation.Models.Sheets
{
    /// <summary>
    /// An implementation of IEnumable&lt;ISheet&gt; base on the object having sub-sheets represented in public properties, such as implementation of ISheet,
    /// or classed decorated with ColumnAttribute, or finally public properties looking like simple data columns, 
    /// </summary>
    internal class PropertySheets : AbstractBaseSheets
    {
        private static Dictionary<Type, SheetMeta[]> propertyCache = new Dictionary<Type, SheetMeta[]>();

        public override void Populate(object instance)
        {
            var metas = propertyCache[instance.GetType()];
            var sheetResolver = new Kipon.Excel.WriterImplementation.Factories.SheetResolver();
            foreach (var meta in metas)
            {
                var nextSheet = meta.Property.GetValue(instance);
                if (nextSheet != null)
                {
                    var sheet = sheetResolver.Resolve(nextSheet);
                    var propertySheet = sheet as Kipon.Excel.WriterImplementation.Models.Sheet.AbstractBaseSheet;
                    if (propertySheet != null)
                    {
                        propertySheet.Title = meta.title;
                    }
                    this.Add(sheet);
                }
            }
        }

        /// <summary>
        /// Populate metadata for attributes where the class properties is decorated with SheetAttributes
        /// </summary>
        /// <param name="properties"></param>
        internal void AddSheetAttributeDecoratedProperties(Type type, System.Reflection.PropertyInfo[] properties)
        {
            var sheetMetas = new List<SheetMeta>();

            var ix = 0;
            foreach (var prop in properties)
            {
                var order = ix++;

                var attr = (Kipon.Excel.Attributes.SheetAttribute)prop.GetCustomAttributes(typeof(Kipon.Excel.Attributes.SheetAttribute), true).First();
                var title = attr.Title;
                if (string.IsNullOrEmpty(title))
                {
                    title = prop.Name;
                }
                sheetMetas.Add(new SheetMeta { title = title, order = order, Property = prop });
            }
            propertyCache[type] = sheetMetas.OrderBy(r => r.order).ToArray();
        }


        /// <summary>
        /// Populate metadata based on duck type sheet properties
        /// </summary>
        /// <param name="type"></param>
        /// <param name="properties"></param>
        internal void AddUndecoratedProperties(Type type, System.Reflection.PropertyInfo[] properties)
        {
            var sheetMetas = new List<SheetMeta>();

            var ix = 0;
            foreach (var prop in properties)
            {
                var order = ix++;
                var title = prop.Name;
                sheetMetas.Add(new SheetMeta { title = title, order = order, Property = prop });
            }
            propertyCache[type] = sheetMetas.ToArray();
        }

        private class SheetMeta
        {
            internal string title { get; set; }
            internal int order { get; set; }
            internal System.Reflection.PropertyInfo Property { get; set; }
        }
    }
}
