using System;

namespace Nerosoft.Powersheet;

/// <summary>
/// 指定的表格名称在文档中不存在时抛出此异常
/// </summary>
public class SheetNotFoundException : Exception
{
    /// <summary>
    /// 初始化<see cref="SheetNotFoundException"/>的实例。
    /// </summary>
    /// <param name="name">异常表格名称</param>
    public SheetNotFoundException(string name)
        : base($"The worksheet '{name}' does not exist in workbook.")
    {
        Name = name;
    }

    /// <summary>
    /// 初始化<see cref="SheetNotFoundException"/>的实例。
    /// </summary>
    /// <param name="name">异常表格名称</param>
    /// <param name="message">错误消息</param>
    public SheetNotFoundException(string name, string message)
        : base(message)
    {
        Name = name;
    }

    /// <summary>
    /// 获取异常表格名称
    /// </summary>
    public string Name { get; }
}