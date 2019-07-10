using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Kipon.Excel.Implementation.Serialization;
using Kipon.Excel.Types;

namespace Kipon.Excel.Styles
{
    internal class DefaultStylesheet : Stylesheet, Kipon.Excel.Implementation.Serialization.IStyleResolver
    {
        #region internal statics
        internal static readonly UInt32Value BOLD_STYLE_INDEX = (UInt32Value)4;

        internal static readonly UInt32Value ODD_STYLE_INDEX_LOCKED = (UInt32Value)3;
        internal static readonly UInt32Value EVEN_STYLE_INDEX_LOCKED = (UInt32Value)2;


        internal static readonly UInt32Value ODD_STYLE_INDEX_UNLOCKED = (UInt32Value)1;
        internal static readonly UInt32Value EVEN_STYLE_INDEX_UNLOCKED = (UInt32Value)0;
        #endregion

        #region the stylesheet
        public DefaultStylesheet()
        {
            Fonts fonts = new Fonts() { Count = (UInt32Value)2U };
            Font normalFont = new Font();

            Font boldFont = new Font();
            boldFont.Append(new Bold());
            Color color = new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF" } };
            boldFont.Append(color);


            fonts.Append(normalFont);
            fonts.Append(boldFont);

            Fills fills = new Fills() { Count = (UInt32Value)4U };

            // light blue background
            Fill fill3 = new Fill();
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "DCE6F1" } };
            patternFill3.Append(foregroundColor1);
            fill3.Append(patternFill3);

            // dark blue
            Fill fill4 = new Fill();
            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "4F81BD" } };

            patternFill4.Append(foregroundColor2);
            fill4.Append(patternFill4);

            fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.None })); // required
            fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.Gray125 })); // required
            fills.Append(fill3);
            fills.Append(fill4);

            Borders borders = new Borders() { Count = (UInt32Value)1U };
            borders.Append(new Border());

            CellFormat evenFontUnlocked = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            evenFontUnlocked.Append(new Protection { Locked = false });

            CellFormat oddFontUnlocked = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            oddFontUnlocked.Append(new Protection { Locked = false });

            CellFormat evenFontLocked = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            evenFontLocked.Append(new Protection { Locked = true });

            CellFormat oddFontLocked = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            oddFontLocked.Append(new Protection { Locked = true });


            CellFormat boldFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };


            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)3U };

            cellFormats.Append(evenFontUnlocked);
            cellFormats.Append(oddFontUnlocked);

            cellFormats.Append(evenFontLocked);
            cellFormats.Append(oddFontLocked);

            cellFormats.Append(boldFormat);

            Append(fonts);
            Append(fills);
            Append(borders);
            Append(cellFormats);
        }

        #region impl and hide style resolver
        uint IStyleResolver.Resolve(ISheet sheet, Types.Cell cell)
        {
            if (cell.Row.Value == 0) return BOLD_STYLE_INDEX;
            if (cell.Row.Value % 2 == 0) return EVEN_STYLE_INDEX_UNLOCKED;
            return ODD_STYLE_INDEX_UNLOCKED;
        }
        #endregion
        #endregion
    }
}
