using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Kipon.Excel.WriterImplementation.Serialization;
using Kipon.Excel.Api;
using Kipon.Excel.WriterImplementation.OpenXml.Types;

namespace Kipon.Excel.WriterImplementation.OpenXml.Styles
{
    internal class DefaultStylesheet : Stylesheet, Kipon.Excel.WriterImplementation.Serialization.IStyleResolver
    {

        #region the stylesheet
        public DefaultStylesheet()
        {
            NumberingFormats numberingFormats1 = new NumberingFormats() { Count = (UInt32Value)5U };

            {
                NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)164U, FormatCode = "0" };
                numberingFormats1.Append(numberingFormat1);
            }

            {
                NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)165U, FormatCode = "0.0" };
                numberingFormats1.Append(numberingFormat1);
            }

            {
                NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)166U, FormatCode = "0.000" };
                numberingFormats1.Append(numberingFormat1);
            }

            {
                NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)167U, FormatCode = "0.0000" };
                numberingFormats1.Append(numberingFormat1);
            }

            {
                NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)168U, FormatCode = "0.00000" };
                numberingFormats1.Append(numberingFormat1);
            }

            /*
            for (var i = (UInt32Value)165U; i <= (UInt32Value)169U; i++)
            {
                var x = (int)i.Value;
                var decs = "0".PadLeft(x, '0');
                numberingFormat1 = new NumberingFormat() { NumberFormatId = i, FormatCode = $"0.{ decs }" };
                numberingFormats1.Append(numberingFormat1);
            }
            */

            Fonts fonts = new Fonts() { Count = (UInt32Value)2U };
            Font normalFont = new Font();

            Font boldFont = new Font();
            boldFont.Append(new Bold());
            Color color = new Color() { Rgb = HexBinaryValue.FromString("FFFFFF")  };
            boldFont.Append(color);

            fonts.Append(normalFont);
            fonts.Append(boldFont);

            Fills fills = new Fills() { Count = (UInt32Value)4U };
            fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.None })); // required
            fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.Gray125 })); // required

            // light blue background
            {
                Fill fill3 = new Fill();
                PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = HexBinaryValue.FromString("DCE6F1") };
                patternFill3.Append(foregroundColor1);
                fill3.Append(patternFill3);
                fills.Append(fill3);
            }

            // dark blue
            {
                Fill fill4 = new Fill();
                PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = HexBinaryValue.FromString("4F81BD") };
                patternFill4.Append(foregroundColor2);

                fill4.Append(patternFill4);
                fills.Append(fill4);
            }

            Borders borders = new Borders() { Count = (UInt32Value)1U };
            borders.Append(new Border());

            CellFormats cellFormats = new CellFormats() { Count = 34 };
            cellFormats.Append(new CellFormat());

            void AddStyles(UInt32Value lowNumberFormat)
            {
                CellFormat evenFontUnlocked = new CellFormat() { NumberFormatId = lowNumberFormat, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
                evenFontUnlocked.Append(new Protection { Locked = false });

                CellFormat oddFontUnlocked = new CellFormat() { NumberFormatId = lowNumberFormat, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U };
                oddFontUnlocked.Append(new Protection { Locked = false });

                CellFormat evenFontLocked = new CellFormat() { NumberFormatId = lowNumberFormat, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
                evenFontLocked.Append(new Protection { Locked = true });

                CellFormat oddFontLocked = new CellFormat() { NumberFormatId = lowNumberFormat, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U };
                oddFontLocked.Append(new Protection { Locked = true });

                cellFormats.Append(evenFontUnlocked);
                cellFormats.Append(oddFontUnlocked);

                cellFormats.Append(evenFontLocked);
                cellFormats.Append(oddFontLocked);
            }

            AddStyles((UInt32Value)0U);

            CellFormat boldFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, ApplyFont = true };
            boldFormat.Append(new Protection { Locked = true });

            CellFormat evenDateUnlocked = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = true };
            evenDateUnlocked.Append(new Protection { Locked = false });

            CellFormat oddDateUnlocked = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, ApplyNumberFormat = true };
            oddDateUnlocked.Append(new Protection { Locked = false });

            CellFormat evenDateLocked = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = true };
            evenDateLocked.Append(new Protection { Locked = true });

            CellFormat oddDateLocked = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, ApplyNumberFormat = true };
            oddDateLocked.Append(new Protection { Locked = true });

            cellFormats.Append(boldFormat);
            cellFormats.Append(evenDateUnlocked);
            cellFormats.Append(oddDateUnlocked);
            cellFormats.Append(evenDateLocked);
            cellFormats.Append(oddDateLocked);

            AddStyles((UInt32Value)164U); // 0 decimals
            AddStyles((UInt32Value)165U); // 1 decimals
            AddStyles((UInt32Value)2U); // 2 decimals
            AddStyles((UInt32Value)166U); // 3 decimals
            AddStyles((UInt32Value)167U); // 4 decimals
            AddStyles((UInt32Value)168U); // 5 decimals

            /*
            for (var i = (UInt32Value)164U; i <= (UInt32Value)169U; i++)
            {
                AddStyles(i, i);
            }
            */


            Append(numberingFormats1);
            Append(fonts);
            Append(fills);
            Append(borders);
            Append(cellFormats);
        }

        #region impl and hide style resolver
        uint IStyleResolver.Resolve(Kipon.Excel.Api.ICell cell)
        {
            var row = cell.Coordinate.Point.Last();
            return Kipon.Excel.Api.Style.Styles.Resolve(cell, row % 2 == 0);
        }
        #endregion
        #endregion
    }
}
