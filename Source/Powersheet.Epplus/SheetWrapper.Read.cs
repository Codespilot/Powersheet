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
    /// <inherited/>
    public override async Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, int sheetIndex, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var excel = new ExcelPackage(stream);
            var sheet = excel.Workbook.Worksheets[sheetIndex];

            return Read(sheet, firstRowNumber, columnNumber, valueConvert);
        }, cancellationToken);
    }

    /// <inherited/>
    public override async Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var excel = new ExcelPackage(stream);
            var sheet = GetSheet(excel, sheetName);

            return Read(sheet, firstRowNumber, columnNumber, valueConvert);
        }, cancellationToken);
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

    /// <summary>
    /// 读取表格内容并按行转换为对象集合
    /// </summary>
    /// <typeparam name="T">数据对象类型</typeparam>
    /// <param name="stream"></param>
    /// <param name="options"></param>
    /// <param name="mapperAction"></param>
    /// <param name="itemAction"></param>
    /// <param name="valueAction"></param>
    /// <param name="sheetIndex"></param>
    /// <returns></returns>
    protected override List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, int sheetIndex = 0)
    {
        if (sheetIndex < 0)
        {
            throw new IndexOutOfRangeException("The sheet index could not less than 0.");
        }

        var excel = new ExcelPackage(stream);
        if (sheetIndex > excel.Workbook.Worksheets.Count - 1)
        {
            throw new IndexOutOfRangeException($"The sheet index could not larger than or equals the workbook sheet count ({excel.Workbook.Worksheets.Count}).");
        }

        var sheet = excel.Workbook.Worksheets[sheetIndex];

        return Read(sheet, options, mapperAction, itemAction, valueAction);
    }

    /// <summary>
    /// 读取表格内容并按行转换为对象集合
    /// </summary>
    /// <typeparam name="T">数据对象类型</typeparam>
    /// <param name="stream"></param>
    /// <param name="options"></param>
    /// <param name="mapperAction"></param>
    /// <param name="itemAction"></param>
    /// <param name="valueAction"></param>
    /// <param name="sheetName"></param>
    /// <returns></returns>
    protected override List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, string sheetName)
    {
        var excel = new ExcelPackage(stream);
        var sheet = GetSheet(excel, sheetName);

        return Read(sheet, options, mapperAction, itemAction, valueAction);
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

    private static ExcelWorksheet GetSheet(ExcelPackage excel, string sheetName)
    {
        ExcelWorksheet sheet;
        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            var names = GetSheetNames();
            if (names.All(t => t.Equals(sheetName)))
            {
                throw new SheetNotFoundException(sheetName, $"The workbook does not contains a sheet named '{sheetName}'");
            }

            sheet = excel.Workbook.Worksheets[sheetName];
        }
        else
        {
            sheet = excel.Workbook.Worksheets[0];
        }

        return sheet;

        IEnumerable<string> GetSheetNames()
        {
            for (var index = 0; index < excel.Workbook.Worksheets.Count; index++)
            {
                yield return excel.Workbook.Worksheets[index].Name;
            }
        }
    }
}