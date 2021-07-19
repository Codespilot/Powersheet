using System;
using System.Globalization;

namespace Nerosoft.Powersheet
{
    public class SheetColumnMapProfile
    {
        public SheetColumnMapProfile()
        { }

        public SheetColumnMapProfile(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentNullException(nameof(name));
            }

            Name = name;
        }

        public SheetColumnMapProfile(string name, string columnName)
            : this(name)
        {
            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentNullException(nameof(columnName));
            }
            ColumnName = columnName;
        }

        public SheetColumnMapProfile(string name, string columnName, Func<object, CultureInfo, object> valueConvert) 
            : this(name, columnName)
        {
            ValueConvert = valueConvert;
        }

        public SheetColumnMapProfile(string name, string columnName, ICellValueConverter valueConverter) 
            : this(name, columnName)
        {
            ValueConverter = valueConverter;
        }

        /// <summary>
        /// 获取或设置对属性或DataTable列的名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置表格对应列的名称
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// 获取或设置一个值标识被标记设置的属性或DataTable列是否忽略。
        /// 值为<c>true</c>时其他属性的设置均会失效。
        /// </summary>
        public bool Ignore { get; set; }

        /// <summary>
        /// 获取或设置一个值用于在写入写入表格的时候指定列的顺序。
        /// 值越小列越靠前。
        /// </summary>
        public int Order { get; set; }

        /// <summary>
        /// 获取或设置单元格内容转换委托方法。
        /// </summary>
        public Func<object, CultureInfo, object> ValueConvert { get; set; }

        /// <summary>
        /// 获取或设置单元格内容转换器。
        /// 如果同时设置列<see cref="ValueConvert"/>，此设置无效。
        /// </summary>
        public ICellValueConverter ValueConverter { get; set; }
    }
}
