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

        private void Convert<T>(T target, Models.Spreadsheet spreadsheet)
        {

            var targetType = typeof(T);


            foreach (var sheet in spreadsheet.Sheets)
            {
                var rows = (from c in sheet.Cells
                            group c by c.Coordinate.Point.Last() into row
                            select new
                            {
                                Row = row.Key,
                                Cells = row.ToArray()
                            }).OrderBy(r => r.Row).ToArray();

                if (rows.Length < 1)
                {
                    continue;
                }

                System.Collections.IList list = null;
                Type elementType = null;
                if (targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(List<>))
                {
                    elementType = targetType.GetGenericArguments().First();
                    list = (System.Collections.IList)target;
                }
                else
                {
                }

                var headers = rows.First().Cells.Where(r => r.Value != null).Select(r => r.Value.ToString()).ToArray();

                foreach (var row in rows.Skip(1))
                {
                    var next = elementType.GetConstructor(new Type[0]).Invoke(new object[0]);
                    list.Add(next);
                }

                foreach (var row in rows)
                {
                    Console.Write(row.Row.ToString().PadLeft(3, ' ') + ": ");
                    foreach (var cell in row.Cells)
                    {
                        if (cell.Value == null)
                        {
                            Console.Write("null");
                        } else
                        {
                            Console.Write(cell.Value.ToString());
                        }
                        Console.Write("\t");
                    }
                    Console.WriteLine("");
                }
                Console.WriteLine("===================================");
            }
        }
    }
}
