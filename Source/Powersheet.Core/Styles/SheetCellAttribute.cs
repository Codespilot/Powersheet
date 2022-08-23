using System;
using System.Drawing;

namespace Nerosoft.Powersheet;

/// <summary>
/// 单元格样式
/// </summary>
public abstract class SheetCellAttribute : Attribute, ISheetStyleAttribute
{
    /// <summary>
    /// 获取或设置单元格背景颜色
    /// </summary>
    public Color FillColor { get; set; } = Color.Empty;

    /// <summary>
    /// 获取或设置文字是否折行显示
    /// </summary>
    public bool WrapText { get; set; }

    /// <summary>
    /// 获取或设置文字水平对齐方式
    /// </summary>
    public HorizontalAlignment HorizontalAlignment { get; set; } = HorizontalAlignment.Left;

    /// <summary>
    /// 获取或设置文字垂直对齐方式
    /// </summary>
    public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Top;

    public void SetStyle(CellStyle style)
    {
        style ??= new CellStyle();
        style.FillColor = FillColor;
        style.WrapText = WrapText;
        style.HorizontalAlignment = HorizontalAlignment;
        style.VerticalAlignment = VerticalAlignment;
    }
}

/// <summary>
/// 表头单元格样式
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class SheetHeaderCellAttribute : SheetCellAttribute
{
}

/// <summary>
/// 数据单元格样式
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class SheetBodyCellAttribute : SheetCellAttribute
{
}