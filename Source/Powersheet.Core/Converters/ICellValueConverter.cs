using System.Globalization;

namespace Nerosoft.Powersheet
{
    /// <summary>
    /// 单元格内容转换器契约接口
    /// </summary>
    public interface ICellValueConverter
    {
        /// <summary>
        /// 将单元格的内容转换为对象属性值或DataTable列的值。
        /// </summary>
        /// <param name="value">单元格内容</param>
        /// <param name="culture">语言文化配置，默认为<see cref="CultureInfo.CurrentCulture"/>。</param>
        /// <returns></returns>
        object Convert(object value, CultureInfo culture);

        /// <summary>
        /// 将对象属性值或或DataTable列的值转换为单元格内容。
        /// </summary>
        /// <param name="value">对象属性值或或DataTable列值</param>
        /// <param name="culture">语言文化配置，默认为<see cref="CultureInfo.CurrentCulture"/>。</param>
        /// <returns></returns>
        object ConvertBack(object value, CultureInfo culture);
    }
}
