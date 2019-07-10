using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Api;
using Kipon.Excel.Api.Globalization;

namespace Kipon.Excel.Implementation.Serialization
{
    internal class InternalExcelSerializer
    {
        private ISpreadsheet _spreadsheet;
        private IStyleResolver _styleResolver;
        private IDataTypeResolver _datatypeResolver;
        private ILocalization _localization;

        private NumberFormatInfo valueNumberFormatInfo = new NumberFormatInfo() { NumberDecimalSeparator = ".", NumberGroupSeparator = string.Empty };


        internal InternalExcelSerializer(ISpreadsheet spreadsheet, IDataTypeResolver datatyperesolver, ILocalization localization = null)
        {
            this._spreadsheet = spreadsheet;
            this._datatypeResolver = datatyperesolver;

            if (localization != null)
            {
                this._localization = localization;
            } else
            {
                this._localization = new Kipon.Excel.Implementation.DefaultLocalization();
            }
        }

        internal void Serialize(System.IO.Stream stream)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                createParts(document);
            }
        }

        private void createParts(SpreadsheetDocument document)
        {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            generateWorkbookPartContent(workbookPart);

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId" + (this._spreadsheet.Sheets.Count() + 1).ToString());
            workbookStylesPart.Stylesheet = new Kipon.Excel.Styles.DefaultStylesheet();
            this._styleResolver = (IStyleResolver)workbookStylesPart.Stylesheet;

            var ix = 1;
            foreach (var sheet in _spreadsheet.Sheets)
            {
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId" + ix.ToString());
                generateWorksheetPartContent(worksheetPart, sheet);
                ix++;
            }

        }

        private void generateWorkbookPartContent(WorkbookPart workbookPart)
        {
            Workbook workbook = new Workbook();

            Sheets sheets = new Sheets();

            UInt32 ix = 1;
            foreach (var _sheet in _spreadsheet.Sheets)
            {
                Sheet sheet = new Sheet() { Name = (_spreadsheet.Sheets.First().Title ?? "Sheet " + ix.ToString()).Replace("/", "-"), SheetId = (UInt32Value)ix, Id = "rId" + ix.ToString() };
                sheets.Append(sheet);
                ix++;
            }

            workbook.Append(sheets);
            workbookPart.Workbook = workbook;
        }

        private void generateWorksheetPartContent(WorksheetPart worksheetPart, ISheet sheet)
        {
            Worksheet worksheet = new Worksheet()
            {
                SheetProperties = new SheetProperties()
                {
                    OutlineProperties = new OutlineProperties { SummaryBelow = false },
                }
            };

            SheetData sheetData = new SheetData();

            uint rowix = 0;
            foreach (var dataRow in sheet.Rows)
            {
                Row excelRow = new Row()
                {
                    RowIndex = (UInt32Value)(rowix + 1)
                };

                uint colix = 0;
                foreach (var dataCell in dataRow.Cells)
                {
                    var position = new Api.Types.Cell(colix, rowix);

                    Cell excelCell = new Cell()
                    {
                        CellReference = position.ToString(),
                        StyleIndex = this._styleResolver.Resolve(sheet, position),
                        DataType = this._datatypeResolver.Resolve(sheet, position)
                    };

                    if (dataCell.Value != null)
                    {
                        switch (excelCell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                {
                                    throw new NotSupportedException("for now, shared string is not supported.");
                                }
                            case CellValues.Boolean:
                                {
                                    if (dataCell.Value is bool)
                                    {
                                        var excelValue = new CellValue(this._localization.ToLocal(dataCell.Value).ToString());
                                        excelCell.Append(excelValue);
                                        break;
                                    }
                                    throw new InvalidCastException(dataCell.Value.GetType().FullName + " cannot be converted to boolean");
                                }
                            case CellValues.Date:
                                {
                                    if (dataCell.Value is DateTime)
                                    {
                                        var excelValue = new CellValue((DateTime)dataCell.Value);
                                        excelCell.Append(excelValue);
                                        break;
                                    }
                                    throw new InvalidCastException(dataCell.Value.GetType().FullName + " cannot be converted to DateTime");
                                }
                            case CellValues.Number:
                                {
                                    if (dataCell.Value is byte || dataCell.Value is short || dataCell.Value is int || dataCell.Value is long)
                                    {
                                        var excelValue = new CellValue(dataCell.Value.ToString());
                                        excelCell.Append(excelValue);
                                        break;
                                    }

                                    if (dataCell.Value is double || dataCell.Value is decimal || dataCell.Value is float)
                                    {
                                        var decimals = this._datatypeResolver.NumberOfDecimals(sheet, position);
                                        var format = "{ 0:0.####################}";
                                        if (decimals != null)
                                        {
                                            var zeros = string.Empty.PadLeft(decimals.Value, '0');
                                            format = format.Replace("####################", zeros);
                                        }
                                        var excelValue = new CellValue(string.Format(valueNumberFormatInfo, format, dataCell.Value));
                                        excelCell.Append(excelValue);
                                        break;
                                    }
                                    throw new InvalidCastException(dataCell.Value.GetType().FullName + " cannot be converted to number");
                                }
                            case CellValues.String:
                                {
                                    var excelValue = new CellValue(dataCell.Value.ToString());
                                    excelCell.Append(excelValue);
                                    break;
                                }
                            default:
                                {
                                    throw new NotSupportedException("type " + excelCell.DataType.Value + " not supported.");
                                }
                        }
                    }
                    excelRow.Append(excelCell);
                    colix++;
                }

                sheetData.Append(excelRow);
                rowix++;
            }

            worksheet.Append(sheetData);
            worksheetPart.Worksheet = worksheet;
        }
    }
}
