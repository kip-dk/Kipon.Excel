using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Kipon.Excel.ReaderImplementation.OpenXml
{
    internal class OpenXmlReader
    {
        private Api.ExcelStream context;

        internal OpenXmlReader(Api.ExcelStream context)
        {
            this.context = context;
        }

        internal Models.Spreadsheet Parse(System.IO.Stream stream)
        {
            var result = new Models.Spreadsheet();

            using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                var sharedText = new Dictionary<int, string>();
                if (workbookPart.SharedStringTablePart != null && workbookPart.SharedStringTablePart.SharedStringTable != null)
                {
                    var shi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>();

                    var ix = 0;
                    foreach (var e in shi)
                    {
                        if (e.Text != null && e.Text.Text != null)
                        {
                            sharedText[ix] = e.Text.Text.Trim();
                        }
                        else
                        {
                            if (e.InnerText != null)
                            {
                                sharedText[ix] = e.InnerText.Trim();
                            }
                        }
                        ix++;
                    }
                }

                Dictionary<string, Kipon.Excel.ReaderImplementation.Models.Sheet> sheets = new Dictionary<string, Kipon.Excel.ReaderImplementation.Models.Sheet>();
                foreach (Sheet excelSheet in workbookPart.Workbook.Sheets)
                {
                    var sheet = result.Add(excelSheet.Name);
                    sheets.Add(excelSheet.Id, sheet);
                }

                foreach (var key in sheets.Keys)
                {
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(key);
                    var sheet = sheets[key];
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            if (c.CellValue == null)
                            {
                                continue;
                            }

                            var dataType = c.DataType != null ? c.DataType.Value : CellValues.String;

                            {
                                Kipon.Excel.WriterImplementation.OpenXml.Types.Cell name = c.CellReference.ToString();

                                switch (dataType)
                                {
                                    case CellValues.SharedString:
                                        {
                                            var index = Int32.Parse(c.CellValue.Text);
                                            sheet.Add(name.Column.Index, name.Row.Value, sharedText[index]);
                                            break;
                                        }
                                    case CellValues.Date:
                                        {
                                            var value = DateTime.Parse(c.CellValue.Text);
                                            sheet.Add(name.Column.Index, name.Row.Value, value);
                                            break;
                                        }
                                    case CellValues.Boolean:
                                        {
                                            var txt = c.CellValue.Text != null ? c.CellValue.Text.ToUpper() : "0";
                                            var bol = txt == "1" || txt == "TRUE";
                                            sheet.Add(name.Column.Index, name.Row.Value, bol);
                                            break;
                                        }
                                    case CellValues.InlineString:
                                    case CellValues.String:
                                        {
                                            sheet.Add(name.Column.Index, name.Row.Value, c.CellValue.Text);
                                            break;
                                        }
                                    case CellValues.Number:
                                        {
                                            sheet.Add(name.Column.Index, name.Row.Value, c.CellValue.Text);
                                            break;
                                        }
                                    case CellValues.Error:
                                        {
                                            break;
                                        }
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }
    }
}
