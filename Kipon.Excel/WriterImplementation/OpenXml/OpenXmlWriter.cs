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
using Kipon.Excel.WriterImplementation.Serialization;

namespace Kipon.Excel.WriterImplementation.OpenXml
{
    internal class OpenXmlWriter
    {
        private ISpreadsheet _spreadsheet;
        private IStyleResolver _styleResolver;
        private ILocalization _localization;
        private IDataTypeResolver _datatypeResolver;

        private NumberFormatInfo valueNumberFormatInfo = new NumberFormatInfo() { NumberDecimalSeparator = ".", NumberGroupSeparator = string.Empty };


        internal OpenXmlWriter(ISpreadsheet spreadsheet, ILocalization localization = null)
        {
            this._spreadsheet = spreadsheet;

            if (localization != null)
            {
                this._localization = localization;
            } else
            {
                this._localization = new Kipon.Excel.WriterImplementation.DefaultLocalization();
            }

            this._datatypeResolver = new Kipon.Excel.WriterImplementation.OpenXml.DataTypeResolver();
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
            workbookStylesPart.Stylesheet = new Kipon.Excel.WriterImplementation.OpenXml.Styles.DefaultStylesheet();
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
                Sheet sheet = new Sheet() { Name = (_sheet.Title ?? "Sheet " + ix.ToString()).Replace("/", "-"), SheetId = (UInt32Value)ix, Id = "rId" + ix.ToString() };
                sheets.Append(sheet);
                ix++;
            }

            workbook.Append(sheets);
            workbookPart.Workbook = workbook;
        }

        private void generateWorksheetPartContent(WorksheetPart worksheetPart, ISheet sheet)
        {
            var tmpRows = (from c in sheet.Cells
                        select new
                        {
                            OpenxmlCell = WriterImplementation.OpenXml.Types.Cell.getCell(System.Convert.ToUInt32(c.Coordinate.Point.First()), System.Convert.ToUInt32(c.Coordinate.Point.Last())),
                            Value = c.Value,
                            Cell = c
                        }).ToArray();

            var rows = (from r in tmpRows
                        group r by r.OpenxmlCell.Row.Value into grp
                        select new
                        {
                            Index = grp.Key,
                            Cells = grp.ToArray()
                        }).ToArray();

            Worksheet worksheet = new Worksheet()
            {
                SheetProperties = new SheetProperties()
                {
                    OutlineProperties = new OutlineProperties { SummaryBelow = false },
                }
            };

            var validates = this.generateColumnDefinitions(worksheet, sheet);

            SheetData sheetData = new SheetData();

            uint rowix = 0;
            foreach (var dataRow in rows)
            {
                Row excelRow = new Row()
                {
                    RowIndex = (UInt32Value)(rowix + 1)
                };

                uint colix = 0;
                foreach (var dataCell in dataRow.Cells)
                {
                    var position = WriterImplementation.OpenXml.Types.Cell.getCell(colix, rowix);

                    Cell excelCell = new Cell()
                    {
                        CellReference = position.ToString(),
                        StyleIndex = this._styleResolver.Resolve(dataCell.Cell),
                        DataType = _datatypeResolver.Resolve(dataCell.Cell)
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
                                        var txt = (bool)dataCell.Value ? this._localization.True : this._localization.False;
                                        var excelValue = new CellValue(txt);
                                        excelCell.Append(excelValue);
                                        break;
                                    }
                                    throw new InvalidCastException(dataCell.Value.GetType().FullName + " cannot be converted to boolean");
                                }
                            case CellValues.Date:
                                {
                                    excelCell.DataType = CellValues.Number;

                                    if (dataCell.Value is DateTime)
                                    {
                                        var dateValue = ((DateTime)dataCell.Value).ToOADate();
                                        var excelValue = new CellValue(dateValue.ToString(CultureInfo.InvariantCulture));
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
                                        var decimals = _datatypeResolver.NumberOfDecimals(dataCell.Cell);
                                        var format = "{0:0.####################}";
                                        if (decimals != null)
                                        {
                                            var zeros = string.Empty.PadLeft(decimals.Value, '0');
                                            format = format.Replace("####################", zeros);
                                        }
                                        var valueString = string.Format(valueNumberFormatInfo, format, dataCell.Value);
                                        var excelValue = new CellValue(valueString);
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

                    if (excelCell.DataType != null && excelCell.DataType.Value == CellValues.Date)
                    {
                        excelCell.DataType = CellValues.Number;
                    }

                    if (excelCell.DataType != null && excelCell.DataType.Value == CellValues.Boolean)
                    {
                        excelCell.DataType = CellValues.String;
                    }
                    excelRow.Append(excelCell);
                    colix++;
                }

                sheetData.Append(excelRow);
                rowix++;
            }

            worksheet.Append(sheetData);

            var sheetProtection = new SheetProtection()
            {
                Sheet = true,
                Objects = true,
                Scenarios = true,
                FormatCells = false,
                FormatRows = false,
                FormatColumns = false,
                InsertRows = false,
                DeleteRows = false,
                Sort = false
                /*
                AlgorithmName = "SHA-512",
                HashValue = "LObiQjq3cB5dG7o09BKWcaYSv5yZM6zvIZNrp0uqAttdaXnL7yLEr6OowSX4luXNDI5eMjtAaoF6qIYbIVKe9w==",
                SaltValue = "AvT9iRlToC0sfmW2ocVDDQ==",
                SpinCount = 100000
                */
            };
            worksheet.Append(sheetProtection);


            if (validates != null)
            {
                worksheet.Append(validates);
            }

            worksheetPart.Worksheet = worksheet;
        }

        private DataValidations generateColumnDefinitions(Worksheet worksheet, ISheet sheet)
        {
            var columnDef = sheet as Kipon.Excel.WriterImplementation.Serialization.IColumns;
            if (columnDef != null)
            {
                var sheetColumns = columnDef.Columns != null ? columnDef.Columns.ToArray() : null;
                if (sheetColumns != null && sheetColumns.Length > 0)
                {
                    var columns = new Columns();
                    var needValidation = (from co in sheetColumns where co.MaxLength != null || (co.OptionSetValue != null && co.OptionSetValue.Length > 0) select co).Count();
                    var validates = new DataValidations() { Count = System.Convert.ToUInt32(needValidation) };

                    for (int i = 0; i < sheetColumns.Length; i++)
                    {
                        var sheetColumn = sheetColumns[i];
                        var col = new Column { Min = System.Convert.ToUInt32(i + 1), Max = System.Convert.ToUInt32(i + 1) };
                        if (sheetColumn.Width != null)
                        {
                            col.Width = sheetColumn.Width;
                        } else
                        {
                            col.Width = 12d;
                        }

                        if (sheetColumn.Hidden != null && sheetColumn.Hidden.Value)
                        {
                            col.Hidden = true;
                        }
                        columns.Append(col);

                        #region append maxlength validations
                        if (sheetColumn.MaxLength != null && sheetColumn.MaxLength > 0)
                        {
                            var columnName = Kipon.Excel.WriterImplementation.OpenXml.Types.Column.getColumn(i).Value;
                            DataValidation dv = new DataValidation()
                            {
                                Type = DataValidationValues.TextLength,
                                Operator = DataValidationOperatorValues.LessThanOrEqual,
                                AllowBlank = true
                                ,
                                ShowInputMessage = true,
                                ShowErrorMessage = true,
                                ErrorTitle = this._localization.LengthExceededErrorTitle,
                                Error = this._localization.LengthExceededError(sheetColumn.MaxLength.Value),
                                PromptTitle = this._localization.LengthExceededPromptTitle,
                                Prompt = this._localization.LengthExceededPrompt(sheetColumn.MaxLength.Value),
                                SequenceOfReferences = new ListValue<StringValue>() { InnerText = $"{columnName}2:{columnName}1048576" },
                                Formula1 = new Formula1($"{sheetColumn.MaxLength.Value.ToString()}")
                            };
                            validates.Append(dv);
                            continue;
                        }
                        #endregion

                        #region append enum list validations
                        if (sheetColumn.OptionSetValue != null && sheetColumn.OptionSetValue.Length > 0)
                        {
                            var columnName = Kipon.Excel.WriterImplementation.OpenXml.Types.Column.getColumn(i).Value;
                            DataValidation dv = new DataValidation()
                            {
                                Type = DataValidationValues.List,
                                AllowBlank = true,
                                ShowInputMessage = false,
                                ShowErrorMessage = false,
                                SequenceOfReferences = new ListValue<StringValue>() { InnerText = $"{columnName}2:{columnName}1048576" },
                                Formula1 = new Formula1("\"" + string.Join(",", sheetColumn.OptionSetValue) + "\"")
                            };
                            validates.Append(dv);
                            continue;
                        }
                        #endregion
                    }
                    worksheet.Append(columns);

                    if (needValidation > 0)
                    {
                        return validates;
                    }
                }
            }
            return null;
        }
    }
}
