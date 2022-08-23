using System;
using System.Drawing;

namespace Nerosoft.Powersheet
{
    /// <summary>
    /// 表格样式标记
    /// </summary>
    public abstract class SheetStyleAttribute : Attribute
    {
        private readonly CellStyle _style = new();

        /// <summary>
        /// 获取或设置字体大小，单位px
        /// </summary>
        public virtual double FontSize
        {
            get => _style.FontSize;
            set => _style.FontSize = value;
        }

        /// <summary>
        /// 获取或设置字体名称
        /// </summary>
        public virtual string FontName
        {
            get => _style.FontName;
            set => _style.FontName = value;
        }

        /// <summary>
        /// 获取或设置文字是否加粗
        /// </summary>
        public virtual bool Bold
        {
            get => _style.Bold;
            set => _style.Bold = value;
        }

        /// <summary>
        /// 获取或设置文字是否是斜体
        /// </summary>
        public virtual bool Italic
        {
            get => _style.Italic;
            set => _style.Italic = value;
        }

        /// <summary>
        /// 获取或设置字体颜色
        /// </summary>
        public virtual Color FontColor
        {
            get => _style.FontColor;
            set => _style.FontColor = value;
        }

        /// <summary>
        /// 获取或设置单元格背景颜色
        /// </summary>
        public virtual Color FillColor
        {
            get => _style.FillColor;
            set => _style.FontColor = value;
        }

        /// <summary>
        /// 获取或设置文字是否折行显示
        /// </summary>
        public virtual bool WrapText
        {
            get => _style.WrapText;
            set => _style.WrapText = value;
        }

        /// <summary>
        /// 获取或设置文字水平对齐方式
        /// </summary>
        public virtual HorizontalAlignment HorizontalAlignment
        {
            get => _style.HorizontalAlignment;
            set => _style.HorizontalAlignment = value;
        }

        /// <summary>
        /// 获取或设置文字垂直对齐方式
        /// </summary>
        public virtual VerticalAlignment VerticalAlignment
        {
            get => _style.VerticalAlignment;
            set => _style.VerticalAlignment = value;
        }

        /// <summary>
        /// 获取或设置单元格边框样式
        /// </summary>
        public virtual BorderStyle BorderStyle
        {
            get => _style.BorderStyle;
            set => _style.BorderStyle = value;
        }

        /// <summary>
        /// 获取或设置单元格边框颜色
        /// </summary>
        public virtual Color BorderColor
        {
            get => _style.BorderColor;
            set => _style.BorderColor = value;
        }

        /// <summary>
        /// 获取样式
        /// </summary>
        public virtual CellStyle Style => _style;
    }
    
    /// <summary>
    /// 表头单元格样式标记
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class SheetHeaderStyleAttribute : SheetStyleAttribute
    {
    }
    
    /// <summary>
    /// 数据单元格样式标记
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class SheetBodyStyleAttribute : SheetStyleAttribute
    {
    }
}