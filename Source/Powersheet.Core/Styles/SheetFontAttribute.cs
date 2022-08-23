using System;
using System.Drawing;

namespace Nerosoft.Powersheet;

/// <summary>
/// 字体样式
/// </summary>
public abstract class SheetFontAttribute : Attribute, ISheetStyleAttribute
{
    /// <summary>
    /// 获取或设置字体名称
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 获取或设置字体大小
    /// </summary>
    public double Size { get; set; } = 12;

    /// <summary>
    /// 获取或设置文字是否加粗
    /// </summary>
    public bool Bold { get; set; }

    /// <summary>
    /// 获取或设置文字是否是斜体
    /// </summary>
    public bool Italic { get; set; }

    /// <summary>
    /// 获取或设置字体颜色
    /// </summary>
    public Color Color { get; set; }

    public void SetStyle(CellStyle style)
    {
        style ??= new CellStyle();
        style.Bold = Bold;
        style.Italic = Italic;
        style.FontName = Name;
        style.FontSize = Size;
    }
}

/// <summary>
/// 表头字体样式
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class SheetHeaderFontAttribute : SheetFontAttribute
{
}

/// <summary>
/// 数据字体样式
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class SheetBodyFontAttribute : SheetFontAttribute
{
}