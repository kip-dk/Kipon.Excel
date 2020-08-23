using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Extensions.Strings;

namespace Kipon.Excel.ReaderImplementation.Converters
{
    internal class PropertySheetConverter
    {
        private Type elementType;
        private Kipon.Excel.Reflection.PropertySheet sheetMeta;

        public PropertySheetConverter(Type elementType, Kipon.Excel.Reflection.PropertySheet sheetMeta)
        {
            this.elementType = elementType;
            this.sheetMeta = sheetMeta;
        }

        internal System.Collections.IList Convert(Kipon.Excel.Api.ISheet sheet)
        {
            var listType = typeof(List<>);
            var clType = listType.MakeGenericType(this.elementType);

            var instance = (System.Collections.IList)Activator.CreateInstance(clType);

            var headerRow = sheetMeta.HeaderRow - 1;
            var headerColumn = sheetMeta.HeaderColumn - 1;

            var rows = (from c in sheet.Cells
                        where c.Coordinate.Point.Last() >= headerRow
                          && c.Coordinate.Point.First() >= headerColumn
                        group c by c.Coordinate.Point.Last() into r
                        select new
                        {
                            Row = r.Key,
                            Cells = r.ToArray()
                        }).OrderBy(r => r.Row).ToArray();

            if (rows.Length < 1)
            {
                return instance;
            }

            var headers = rows.First();
            Dictionary<int, Kipon.Excel.Reflection.PropertySheet.SheetProperty> columnIndex = new Dictionary<int, Reflection.PropertySheet.SheetProperty>();
            foreach (var header in headers.Cells)
            {
                if (header.Value != null)
                {
                    var property = (from p in sheetMeta.Properties where p.title.ToRelaxedName() == header.Value.ToString().ToRelaxedName() select p).FirstOrDefault();
                    if (property != null)
                    {
                        columnIndex.Add(header.Coordinate.Point.First(), property);
                    }
                }
            }

            var datas = rows.Skip(1);
            foreach (var data in datas)
            {
                var next = Activator.CreateInstance(this.elementType);
                instance.Add(next);
                var cellConverter = new PropertyCellConverter(next);


                foreach (var cell in data.Cells)
                {
                    var ix = cell.Coordinate.Point.First();
                    if (columnIndex.ContainsKey(ix))
                    {
                        cellConverter.Convert(columnIndex[ix], cell);
                    }
                }
            }
            return instance;
        }
    }
}
