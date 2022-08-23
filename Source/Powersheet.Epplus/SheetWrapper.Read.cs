using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Nerosoft.Powersheet.Epplus;

/// <summary>
/// 基于EPPLUS实现表格解析
/// </summary>
public partial class SheetWrapper
{
    public override async Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, IEnumerable<int> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
    {
        CheckSheets(sheets);

        return await Task.Run(() =>
        {
            var result = new List<T>();
            var excel = new ExcelPackage(stream);

            CheckSheets(sheets, excel.Workbook.Worksheets.Count);

            foreach (var sheetIndex in sheets)
            {
                var sheet = excel.Workbook.Worksheets[sheetIndex];

                var items = Read(sheet, firstRowNumber, columnNumber, valueConvert);
                result.AddRange(items);
            }

            return result;
        }, cancellationToken);
    }

    public override async Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, IEnumerable<string> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
    {
        CheckSheets(sheets);

        return await Task.Run(() =>
        {
            var excel = new ExcelPackage(stream);

            CheckSheets(sheets, () => GetSheetNames(excel));

            return sheets.Aggregate(new List<T>(), (current, sheetName) =>
            {
                var sheet = excel.Workbook.Worksheets[sheetName];

                var items = Read(sheet, firstRowNumber, columnNumber, valueConvert);
                current.AddRange(items);
                return current;
            });
        }, cancellationToken);
    }

    /// <summary>
    /// 读取表格内容并按行转换为对象集合
    /// </summary>
    /// <param name="stream"></param>
    /// <param name="options"></param>
    /// <param name="mapperAction"></param>
    /// <param name="itemAction"></param>
    /// <param name="valueAction"></param>
    /// <param name="sheets"></param>
    /// <typeparam name="T">数据对象类型</typeparam>
    /// <returns></returns>
    /// <exception cref="IndexOutOfRangeException"></exception>
    protected override List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, IEnumerable<int> sheets)
    {
        CheckSheets(sheets);

        var excel = new ExcelPackage(stream);

        CheckSheets(sheets, excel.Workbook.Worksheets.Count);

        var result = new List<T>();
        foreach (var sheetIndex in sheets)
        {
            var sheet = excel.Workbook.Worksheets[sheetIndex];

            var items = Read(sheet, options, mapperAction, itemAction, valueAction);
            result.AddRange(items);
        }

        return result;
    }

    /// <summary>
    /// 读取表格内容并按行转换为对象集合
    /// </summary>
    /// <param name="stream"></param>
    /// <param name="options"></param>
    /// <param name="mapperAction"></param>
    /// <param name="itemAction"></param>
    /// <param name="valueAction"></param>
    /// <param name="sheets"></param>
    /// <typeparam name="T">数据对象类型</typeparam>
    /// <returns></returns>
    protected override List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, IEnumerable<string> sheets)
    {
        CheckSheets(sheets);

        var excel = new ExcelPackage(stream);

        CheckSheets(sheets, () => GetSheetNames(excel));
        var result = new List<T>();
        foreach (var sheetName in sheets)
        {
            var sheet = excel.Workbook.Worksheets[sheetName];

            var items = Read(sheet, options, mapperAction, itemAction, valueAction);
            result.AddRange(items);
        }

        return result;
    }

    /// <summary>
    /// 读取表格内容并按行转换为对象集合
    /// </summary>
    /// <typeparam name="T">数据对象类型</typeparam>
    /// <param name="sheet">表格对象</param>
    /// <param name="firstRowNumber"></param>
    /// <param name="columnNumber"></param>
    /// <param name="valueConvert"></param>
    /// <returns></returns>
    protected virtual List<T> Read<T>(ExcelWorksheet sheet, int firstRowNumber, int columnNumber, Func<object, CultureInfo, T> valueConvert)
    {
        if (sheet.Dimension.Columns == 1)
        {
            columnNumber = sheet.Dimension.Start.Column;
        }
        else if (columnNumber < sheet.Dimension.Start.Column || columnNumber > sheet.Dimension.End.Column)
        {
            throw new IndexOutOfRangeException($"Column number '{columnNumber}' was not in the valid column number range of {sheet.Dimension.Start.Column}-{sheet.Dimension.End.Column}.");
        }

        var result = new List<T>();

        for (var rowNumber = firstRowNumber; rowNumber <= sheet.Dimension.Rows; rowNumber++)
        {
            T value;

            if (valueConvert != null)
            {
                value = valueConvert(sheet.Cells[rowNumber, columnNumber].Value, CultureInfo.CurrentCulture);
            }
            else
            {
                value = sheet.Cells[rowNumber, columnNumber].GetValue<T>();
            }

            result.Add(value);
        }

        return result;
    }

    /// <summary>
    /// 读取表格内容并按行转换为对象集合
    /// </summary>
    /// <typeparam name="T">数据对象类型</typeparam>
    /// <param name="sheet">表格对象</param>
    /// <param name="options">配置选项</param>
    /// <param name="mapperAction"></param>
    /// <param name="itemAction"></param>
    /// <param name="valueAction"></param>
    /// <returns></returns>
    protected virtual List<T> Read<T>(ExcelWorksheet sheet, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction)
    {
        if (sheet == null)
        {
            throw new NullReferenceException("The sheet can not be null.");
        }

        options ??= new SheetReadOptions();

        options.Validate();

        var mappers = new Dictionary<int, SheetColumnMapProfile>();

        var result = new List<T>();

        var columnCount = sheet.Dimension.Columns;

        for (var index = options.FirstColumnNumber; index <= columnCount; index++)
        {
            var name = sheet.Cells[options.HeaderRowNumber, index].Text;
            if (string.IsNullOrWhiteSpace(name) || options.IgnoreNames.Contains(name, StringComparer.OrdinalIgnoreCase))
            {
                continue;
            }

            var mapper = options.GetMapProfile(name);
            mapper ??= new SheetColumnMapProfile(name, name);
            mappers.Add(index, mapper);
        }

        mapperAction?.Invoke(mappers);

        for (var rowNumber = options.HeaderRowNumber + 1; rowNumber <= sheet.Dimension.Rows; rowNumber++)
        {
            if (result.Count >= options.RowCount)
            {
                break;
            }

            var item = itemAction();

            for (var columnNumber = options.FirstColumnNumber; columnNumber <= columnCount; columnNumber++)
            {
                if (!mappers.TryGetValue(columnNumber, out var mapper))
                {
                    continue;
                }

                var cellValue = sheet.Cells[rowNumber, columnNumber].Value;
                object value;
                if (mapper.ValueConvert != null)
                {
                    value = mapper.ValueConvert(cellValue, CultureInfo.CurrentCulture);
                }
                else if (mapper.ValueConverter != null)
                {
                    value = mapper.ValueConverter.ConvertCellValue(cellValue, CultureInfo.CurrentCulture);
                }
                else
                {
                    value = cellValue;
                }

                valueAction(item, mapper.Name, value);
            }

            result.Add(item);
        }

        return result;
    }

    private static IEnumerable<string> GetSheetNames(ExcelPackage excel)
    {
        return excel.Workbook.Worksheets.Select(t => t.Name);
    }
}