using System.Drawing;

namespace Nerosoft.Powersheet;

/// <summary>
/// 单元格样式
/// </summary>
public class CellStyle
{
    /// <summary>
    /// 获取或设置字体大小，单位px
    /// </summary>
    public double FontSize { get; set; } = 12;

    /// <summary>
    /// 获取或设置字体名称
    /// </summary>
    public string FontName { get; set; }

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
    public Color FontColor { get; set; } = Color.Empty;

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

    /// <summary>
    /// 获取或设置单元格边框样式
    /// </summary>
    public BorderStyle BorderStyle { get; set; } = BorderStyle.None;

    /// <summary>
    /// 获取或设置单元格边框颜色
    /// </summary>
    public Color BorderColor { get; set; } = Color.Empty;
}