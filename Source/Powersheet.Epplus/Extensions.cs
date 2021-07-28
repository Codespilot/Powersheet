using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Nerosoft.Powersheet.Epplus
{
    /// <summary>
    /// 扩张方法
    /// </summary>
    internal static class Extensions
    {
        internal static void SetCellStyle(this ExcelRange cell, CellStyle style)
        {
            if (style == null)
            {
                return;
            }

            if (!style.FillColor.IsEmpty)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(style.FillColor);
            }

            if (!style.FontColor.IsEmpty)
            {
                cell.Style.Font.Color.SetColor(style.FontColor);
            }

            if (style.BorderStyle != BorderStyle.None)
            {
                cell.Style.Border.Bottom.Style = style.BorderStyle.Convert();
                cell.Style.Border.Top.Style = style.BorderStyle.Convert();
                cell.Style.Border.Left.Style = style.BorderStyle.Convert();
                cell.Style.Border.Right.Style = style.BorderStyle.Convert();

                if (!style.BorderColor.IsEmpty)
                {
                    cell.Style.Border.Bottom.Color.SetColor(style.BorderColor);
                    cell.Style.Border.Top.Color.SetColor(style.BorderColor);
                    cell.Style.Border.Left.Color.SetColor(style.BorderColor);
                    cell.Style.Border.Right.Color.SetColor(style.BorderColor);
                }
            }

            if (!string.IsNullOrWhiteSpace(style.FontName))
            {
                cell.Style.Font.Name = style.FontName;
            }

            if (style.FontSize > 0)
            {
                cell.Style.Font.Size = (float) style.FontSize;
            }

            cell.Style.WrapText = style.WrapText;
            cell.Style.HorizontalAlignment = style.HorizontalAlignment.Convert();
            cell.Style.VerticalAlignment = style.VerticalAlignment.Convert();
            cell.Style.Font.Bold = style.Bold;
            cell.Style.Font.Italic = style.Italic;
        }

        internal static ExcelHorizontalAlignment Convert(this HorizontalAlignment alignment)
        {
            return alignment switch
            {
                HorizontalAlignment.Left => ExcelHorizontalAlignment.Left,
                HorizontalAlignment.Center => ExcelHorizontalAlignment.Center,
                HorizontalAlignment.Right => ExcelHorizontalAlignment.Right,
                _ => throw new ArgumentOutOfRangeException()
            };
        }

        internal static ExcelVerticalAlignment Convert(this VerticalAlignment alignment)
        {
            return alignment switch
            {
                VerticalAlignment.Top => ExcelVerticalAlignment.Top,
                VerticalAlignment.Middle => ExcelVerticalAlignment.Center,
                VerticalAlignment.Bottom => ExcelVerticalAlignment.Bottom,
                _ => throw new ArgumentOutOfRangeException()
            };
        }

        internal static ExcelBorderStyle Convert(this BorderStyle borderStyle)
        {
            var value = borderStyle.ToString();
            return Enum.Parse<ExcelBorderStyle>(value, true);
        }
    }
}