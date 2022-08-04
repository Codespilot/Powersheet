using System;
using NPOI.SS.UserModel;

namespace Nerosoft.Powersheet.Npoi;

/// <summary>
/// 扩展方法
/// </summary>
internal static class Extensions
{
    /// <summary>
    /// 设置单元格样式
    /// </summary>
    /// <param name="cell"></param>
    /// <param name="style"></param>
    /// <param name="cellStyle"></param>
    /// <param name="font"></param>
    internal static void SetCellStyle(this ICell cell, CellStyle style, ICellStyle cellStyle, IFont font)
    {
        if (cell == null || style == null || cellStyle == null)
        {
            return;
        }

        if (style.BorderStyle != BorderStyle.None)
        {
            cellStyle.BorderLeft = style.BorderStyle.Convert();
            cellStyle.BorderBottom = style.BorderStyle.Convert();
            cellStyle.BorderRight = style.BorderStyle.Convert();
            cellStyle.BorderBottom = style.BorderStyle.Convert();

            if (!style.BorderColor.IsEmpty)
            {
                var colorIndex = IndexedColors.ValueOf(style.BorderColor.Name)?.Index ?? IndexedColors.Automatic.Index;
                cellStyle.BottomBorderColor = colorIndex;
                cellStyle.TopBorderColor = colorIndex;
                cellStyle.LeftBorderColor = colorIndex;
                cellStyle.RightBorderColor = colorIndex;
            }
        }

        if (!style.FillColor.IsEmpty)
        {
            cellStyle.FillForegroundColor = IndexedColors.ValueOf(style.FillColor.Name)?.Index ?? IndexedColors.Automatic.Index;
            cellStyle.FillPattern = FillPattern.SolidForeground;
        }

        #region Set font

        if (font != null)
        {
            if (!string.IsNullOrWhiteSpace(style.FontName))
            {
                font.FontName = style.FontName;
            }

            if (style.FontSize > 0)
            {
                font.FontHeightInPoints = style.FontSize;
            }

            font.IsBold = style.Bold;
            font.IsItalic = style.Italic;
            if (!style.FontColor.IsEmpty)
            {
                font.Color = IndexedColors.ValueOf(style.FontColor.Name)?.Index ?? IndexedColors.Automatic.Index;
            }

            cellStyle.SetFont(font);
        }

        #endregion

        cellStyle.WrapText = style.WrapText;
        cellStyle.Alignment = style.HorizontalAlignment.Convert();
        cellStyle.VerticalAlignment = style.VerticalAlignment.Convert();

        cell.CellStyle = cellStyle;
    }

    private static NPOI.SS.UserModel.HorizontalAlignment Convert(this HorizontalAlignment alignment)
    {
        return alignment switch
        {
            HorizontalAlignment.Center => NPOI.SS.UserModel.HorizontalAlignment.Center,
            HorizontalAlignment.Left => NPOI.SS.UserModel.HorizontalAlignment.Left,
            HorizontalAlignment.Right => NPOI.SS.UserModel.HorizontalAlignment.Right,
            _ => throw new ArgumentOutOfRangeException()
        };
    }

    private static NPOI.SS.UserModel.VerticalAlignment Convert(this VerticalAlignment alignment)
    {
        return alignment switch
        {
            VerticalAlignment.Top => NPOI.SS.UserModel.VerticalAlignment.Top,
            VerticalAlignment.Middle => NPOI.SS.UserModel.VerticalAlignment.Center,
            VerticalAlignment.Bottom => NPOI.SS.UserModel.VerticalAlignment.Bottom,
            _ => throw new ArgumentOutOfRangeException()
        };
    }

    private static NPOI.SS.UserModel.BorderStyle Convert(this BorderStyle borderStyle)
    {
        var value = borderStyle.ToString();
        return Enum.Parse<NPOI.SS.UserModel.BorderStyle>(value, true);
    }

    /// <summary>
    /// 设置单元格值
    /// </summary>
    /// <param name="cell"></param>
    /// <param name="value"></param>
    internal static void SetCellValue(this ICell cell, object value)
    {
        switch (value)
        {
            case null:
                return;
            case string cellValue:
                cell.SetCellValue(cellValue);
                break;
            case int:
            case long:
            case decimal:
            case double:
                cell.SetCellValue(value.ToString());
                cell.SetCellType(CellType.Numeric);
                break;
            case DateTime cellValue:
                cell.SetCellValue(cellValue);
                break;
            case bool cellValue:
                cell.SetCellValue(cellValue);
                break;
            default:
                cell.SetCellValue(value.ToString());
                break;
        }
    }
}