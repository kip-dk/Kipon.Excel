using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.ReaderImplementation.Converters
{
    internal class SpreadsheetConverter
    {
        internal T Convert<T>(System.IO.Stream excelStream) where T: class
        {
            var reader = new OpenXml.OpenXmlReader();
            var spreadsheet = reader.Parse(excelStream);

            if (typeof(T).IsAssignableFrom(typeof(Kipon.Excel.Api.ISpreadsheet)))
            {
                return spreadsheet as T;
            }

            var tConstructor = typeof(T).GetConstructor(new Type[0]);
            if (tConstructor == null)
            {
                throw new Kipon.Excel.Exceptions.DefaultConstructorRequiredException(typeof(T));
            }

            var result = (T)tConstructor.Invoke(new object[0]);

            this.Convert(result, spreadsheet);
            return result;
        }

        internal void ConvertInto<T>(T instance, System.IO.Stream excelStream)
        {
            var reader = new OpenXml.OpenXmlReader();
            var spreadsheet = reader.Parse(excelStream);
            this.Convert(instance, spreadsheet);
        }

        private void Convert<T>(T target, Models.Spreadsheet spreadsheet)
        {
            var targetType = typeof(T);

            var isPropertySheets = Kipon.Excel.Reflection.PropertySheets.IsPropertySheets(targetType);

            if (isPropertySheets)
            {
                var sheetsMeta = Kipon.Excel.Reflection.PropertySheets.ForType(targetType);
                foreach (var sheetsPropertyMeta in sheetsMeta.Properties)
                {
                    if (!sheetsPropertyMeta.property.CanWrite)
                    {
                        continue;
                    }

                    var sheet = sheetsPropertyMeta.BestMatch(spreadsheet.Sheets);

                    if (sheet != null)
                    {
                        if (sheetsPropertyMeta.property.PropertyType.IsAssignableFrom(typeof(Kipon.Excel.Api.ISheet)))
                        {
                            sheetsPropertyMeta.property.SetValue(target, sheet);
                            continue;
                        }

                        var isPropertySheet = Kipon.Excel.Reflection.PropertySheet.IsPropertySheet(sheetsPropertyMeta.ElementType);
                        if (isPropertySheet)
                        {
                            var sheetMeta = Kipon.Excel.Reflection.PropertySheet.ForType(sheetsPropertyMeta.ElementType);
                            var converter = new Kipon.Excel.ReaderImplementation.Converters.PropertySheetConverter(sheetsPropertyMeta.ElementType, sheetMeta);
                            var result = converter.Convert(sheet);

                            if (sheetsPropertyMeta.property.PropertyType.IsAssignableFrom(result.GetType()))
                            {
                                sheetsPropertyMeta.property.SetValue(target, result);
                                continue;
                            }

                            if (sheetsPropertyMeta.property.PropertyType.IsArray)
                            {
                                var array = Array.CreateInstance(sheetsPropertyMeta.ElementType, result.Count);
                                for (var i = 0; i<result.Count; i++)
                                {
                                    var v = result[i];
                                    array.SetValue(v, i);
                                }
                                sheetsPropertyMeta.property.SetValue(target, array);
                                continue;
                            }
                        }
                    }
                } 
            }
        }
    }
}
