using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Kipon.Excel.Implementation.Serialization;
using Kipon.Excel.Api;
using Kipon.Excel.Implementation.OpenXml.Types;

namespace Kipon.Excel.Implementation.OpenXml.Styles
{
    internal class DefaultStylesheet : Stylesheet, Kipon.Excel.Implementation.Serialization.IStyleResolver
    {
        #region internal statics

        internal static readonly UInt32Value ODD_DATE_LOCKED = (UInt32Value)9;
        internal static readonly UInt32Value EVEN_DATE_LOCKED = (UInt32Value)8;

        internal static readonly UInt32Value ODD_DATE_UNLOCKED = (UInt32Value)7;
        internal static readonly UInt32Value EVEN_DATE_UNLOCKED = (UInt32Value)6;

        internal static readonly UInt32Value BOLD_STYLE_INDEX = (UInt32Value)5;

        internal static readonly UInt32Value ODD_STYLE_INDEX_LOCKED = (UInt32Value)4;
        internal static readonly UInt32Value EVEN_STYLE_INDEX_LOCKED = (UInt32Value)3;

        internal static readonly UInt32Value ODD_STYLE_INDEX_UNLOCKED = (UInt32Value)2;
        internal static readonly UInt32Value EVEN_STYLE_INDEX_UNLOCKED = (UInt32Value)1;
        #endregion

        #region the stylesheet
        public DefaultStylesheet()
        {
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

            CellFormat evenFontUnlocked = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            evenFontUnlocked.Append(new Protection { Locked = false });

            CellFormat oddFontUnlocked = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U };
            oddFontUnlocked.Append(new Protection { Locked = false });

            CellFormat evenFontLocked = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            evenFontLocked.Append(new Protection { Locked = true });

            CellFormat oddFontLocked = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U };
            oddFontLocked.Append(new Protection { Locked = true });

            CellFormat boldFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, ApplyFont = true };
            boldFormat.Append(new Protection { Locked = true });

            CellFormat evenDateUnlocked = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U,  ApplyNumberFormat = true };
            evenDateUnlocked.Append(new Protection { Locked = false });

            CellFormat oddDateUnlocked = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, ApplyNumberFormat = true };
            oddDateUnlocked.Append(new Protection { Locked = false });

            CellFormat evenDateLocked = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = true };
            evenDateLocked.Append(new Protection { Locked = true });

            CellFormat oddDateLocked = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, ApplyNumberFormat = true };
            oddDateLocked.Append(new Protection { Locked = true });


            CellFormats cellFormats = new CellFormats() { Count = 10 };

            cellFormats.Append(new CellFormat());

            cellFormats.Append(evenFontUnlocked);
            cellFormats.Append(oddFontUnlocked);

            cellFormats.Append(evenFontLocked);
            cellFormats.Append(oddFontLocked);

            cellFormats.Append(boldFormat);

            cellFormats.Append(evenDateUnlocked);
            cellFormats.Append(oddDateUnlocked);

            cellFormats.Append(evenDateLocked);
            cellFormats.Append(oddDateLocked);

            Append(fonts);
            Append(fills);
            Append(borders);
            Append(cellFormats);
        }

        #region impl and hide style resolver
        uint IStyleResolver.Resolve(Kipon.Excel.Api.ICell cell)
        {
            var row = cell.Coordinate.Point.Last();

            if (row == 0) return BOLD_STYLE_INDEX;

            var isreadonly = false;
            var readonlyimpl = cell as Kipon.Excel.Api.Cell.IReadonly;
            if (readonlyimpl != null)
            {
                isreadonly = readonlyimpl.IsReadonly;
            }

            var datatype = cell as Kipon.Excel.Api.Cell.IDataType;
            if (datatype != null && datatype.ValueType != null && datatype.ValueType == typeof(DateTime))
            {
                if (row % 2 == 0 && !isreadonly)
                {
                    return EVEN_DATE_UNLOCKED;
                }

                if (row % 2 == 0 && isreadonly)
                {
                    return EVEN_DATE_LOCKED;
                }

                if (!isreadonly)
                {
                    return ODD_DATE_UNLOCKED;
                }
                return ODD_DATE_LOCKED;
            }

            if (row % 2 == 0 && !isreadonly)
            {
                return EVEN_STYLE_INDEX_UNLOCKED;
            }

            if (row % 2 == 0 && isreadonly)
            {
                return EVEN_STYLE_INDEX_LOCKED;
            }

            if (!isreadonly)
            {
                return ODD_STYLE_INDEX_UNLOCKED;
            }

            return ODD_STYLE_INDEX_LOCKED;
        }
        #endregion
        #endregion
    }
}
