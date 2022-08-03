using System;
using System.Globalization;

namespace Nerosoft.Powersheet
{
    /// <summary>
    /// 单元格内容转换器契约接口
    /// </summary>
    public interface IValueConverter
    {
        /// <summary>
        /// 将单元格的内容转换为对象属性值或DataTable列的值。
        /// </summary>
        /// <param name="value">单元格内容</param>
        /// <param name="culture">语言文化配置，默认为<see cref="CultureInfo.CurrentCulture"/>。</param>
        /// <returns></returns>
        object ConvertCellValue(object value, CultureInfo culture);

        /// <summary>
        /// 将对象属性值或DataTable列的值转换为单元格内容。
        /// </summary>
        /// <param name="value">对象属性值或或DataTable列值</param>
        /// <param name="culture">语言文化配置，默认为<see cref="CultureInfo.CurrentCulture"/>。</param>
        /// <returns></returns>
        object ConvertItemValue(object value, CultureInfo culture);
    }

    internal class EmptyValueConverter<TValue> : IValueConverter<TValue>
    {
        private Func<object, CultureInfo, TValue> _cellValueConverter;
        private Func<TValue, CultureInfo, object> _itemValueConverter;

        public EmptyValueConverter()
        {
        }

        public EmptyValueConverter(Func<object, CultureInfo, TValue> cellValueConverter, Func<TValue, CultureInfo, object> itemValueConverter)
        {
            _cellValueConverter = cellValueConverter;
            _itemValueConverter = itemValueConverter;
        }

        public void SetCellValueConverter(Func<object, CultureInfo, TValue> converter)
        {
            _cellValueConverter = converter;
        }

        public void SetItemValueConverter(Func<TValue, CultureInfo, object> converter)
        {
            _itemValueConverter = converter;
        }

        public TValue ConvertCellValue(object value, CultureInfo culture)
        {
            if (_cellValueConverter == null)
            {
                throw new NotImplementedException();
            }

            return _cellValueConverter(value, culture);
        }

        public object ConvertItemValue(TValue value, CultureInfo culture)
        {
            if (_itemValueConverter == null)
            {
                throw new NotImplementedException();
            }

            return _itemValueConverter;
        }

        object IValueConverter.ConvertItemValue(object value, CultureInfo culture) => ConvertItemValue((TValue)value, culture);
        object IValueConverter.ConvertCellValue(object value, CultureInfo culture) => ConvertCellValue(value, culture);
    }

    /// <summary>
    /// 单元格内容转换器契约接口
    /// </summary>
    /// <typeparam name="TValue"></typeparam>
    public interface IValueConverter<TValue> : IValueConverter
    {
        /// <summary>
        /// 将单元格的内容转换为对象属性值或DataTable列的值。
        /// </summary>
        /// <param name="value">单元格内容</param>
        /// <param name="culture">语言文化配置，默认为<see cref="CultureInfo.CurrentCulture"/>。</param>
        /// <returns></returns>
        new TValue ConvertCellValue(object value, CultureInfo culture);

        /// <summary>
        /// 将对象属性值或DataTable列的值转换为单元格内容。
        /// </summary>
        /// <param name="value">对象属性值或或DataTable列值</param>
        /// <param name="culture">语言文化配置，默认为<see cref="CultureInfo.CurrentCulture"/>。</param>
        /// <returns></returns>
        object ConvertItemValue(TValue value, CultureInfo culture);
    }
}