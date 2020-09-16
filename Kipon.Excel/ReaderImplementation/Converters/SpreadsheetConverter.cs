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

            this.Convert(result, spreadsheet, false);
            return result;
        }

        internal void ConvertInto<T>(T instance, System.IO.Stream excelStream, bool mergeAll)
        {
            var reader = new OpenXml.OpenXmlReader();
            var spreadsheet = reader.Parse(excelStream);
            this.Convert(instance, spreadsheet, mergeAll);
        }

        private void Convert<T>(T target, Models.Spreadsheet spreadsheet, bool mergeAll)
        {
            var targetType = typeof(T);

            {
                Type elementType = null;
                if (targetType.IsArray)
                {
                    elementType = targetType.GetElementType();
                }
                else
                {
                    if (targetType.IsGenericType && typeof(System.Collections.IEnumerable).IsAssignableFrom(targetType))
                    {
                        elementType = targetType.GetGenericArguments()[0];
                    }
                }

                if (elementType != null && Kipon.Excel.Reflection.PropertySheet.IsPropertySheet(elementType))
                {
                    var propertySheetMeta = Kipon.Excel.Reflection.PropertySheet.ForType(elementType);
                    var sheets = propertySheetMeta.AllMatch(spreadsheet.Sheets);
                    if (sheets != null && sheets.Length > 0)
                    {
                        foreach (var sheet in sheets)
                        {
                            var converter = new Kipon.Excel.ReaderImplementation.Converters.PropertySheetConverter(elementType, propertySheetMeta);
                            var result = converter.Convert(sheet);

                            var resultList = target as System.Collections.IList;
                            for (var i = 0; i < result.Count; i++)
                            {
                                resultList.Add(result[i]);
                            }

                            if (!mergeAll)
                            {
                                return;
                            }
                        }
                    }
                }
            }

            {
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

                        var sheet = sheetsPropertyMeta.AllMatch(spreadsheet.Sheets)?.FirstOrDefault();

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
                                    for (var i = 0; i < result.Count; i++)
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

                    return;
                }
            }
        }
    }
}
