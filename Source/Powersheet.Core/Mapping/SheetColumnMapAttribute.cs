using System;

namespace Nerosoft.Powersheet;

/// <summary>
/// 对象属性映射标记
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class SheetColumnMapAttribute : Attribute
{
    /// <summary>
    /// 初始化一个新的 <see cref="SheetColumnMapAttribute"/> 实例
    /// </summary>
    public SheetColumnMapAttribute()
    {
    }

    /// <summary>
    /// 初始化一个新的 <see cref="SheetColumnMapAttribute"/> 实例并传入<see cref="ColumnName"/>属性的值.
    /// </summary>
    /// <param name="name">表格内的列名</param>
    public SheetColumnMapAttribute(string name)
        : this()
    {
        ColumnName = name;
    }

    /// <summary>
    /// 获取表格映射到被标记的属性的列名。
    /// </summary>
    public string ColumnName { get; }

    /// <summary>
    /// 获取或设置一个值标识被标记的此属性是否忽略。
    /// 值为<c>true</c>时其他属性的设置均会失效。
    /// </summary>
    public bool Ignore { get; set; }

    /// <summary>
    /// 获取或设置一个值用于在写入写入表格的时候指定列的顺序。
    /// 值越小列越靠前。
    /// </summary>
    public int Order { get; set; }

    /// <summary>
    /// 获取或设置单元格值转换器类型。
    /// 转换器必须实现 <see cref="IValueConverter"/> 的方法并且包含一个无参构造器。 -- or has non-static public method 'public object ConvertCellValue(object value, Type targetType, CultureInfo culture);'.
    /// 转换器通过无参构造器构造的对象必须能够完整的实现<see cref="IValueConverter"/>的Convert方法。
    /// </summary>
    public Type ValueConverter { get; set; }
}