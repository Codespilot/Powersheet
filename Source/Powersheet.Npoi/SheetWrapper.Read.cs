using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Nerosoft.Powersheet.Npoi;

/// <summary>
/// 
/// </summary>
public partial class SheetWrapper
{
    public override async Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, IEnumerable<int> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
    {
        CheckSheets(sheets);

        return await Task.Run(() =>
        {
            var workbook = new XSSFWorkbook(stream);

            CheckSheets(sheets, workbook.NumberOfSheets);

            return sheets.Aggregate(new List<T>(), (current, sheetIndex) =>
            {
                var sheet = workbook.GetSheetAt(sheetIndex);
                var items = Read(sheet, firstRowNumber, columnNumber, valueConvert);
                current.AddRange(items);
                return current;
            });
        }, cancellationToken);
    }

    public override async Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, IEnumerable<string> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
    {
        CheckSheets(sheets);

        return await Task.Run(() =>
        {
            var workbook = new XSSFWorkbook(stream);

            CheckSheets(sheets, () => GetSheetNames(workbook));

            return sheets.Aggregate(new List<T>(), (current, sheetName) =>
            {
                var sheet = workbook.GetSheet(sheetName);
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
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    /// <exception cref="IndexOutOfRangeException"></exception>
    protected override List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, IEnumerable<int> sheets)
    {
        CheckSheets(sheets);

        var excel = new XSSFWorkbook(stream);

        CheckSheets(sheets, excel.NumberOfSheets);

        return sheets.Aggregate(new List<T>(), (current, sheetIndex) =>
        {
            var sheet = excel.GetSheetAt(sheetIndex);
            var items = Read(sheet, options, mapperAction, itemAction, valueAction);
            current.AddRange(items);
            return current;
        });
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
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    /// <exception cref="SheetNotFoundException">指定的表格名称不存在时抛出此异常。</exception>
    protected override List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, IEnumerable<string> sheets)
    {
        CheckSheets(sheets);

        var workbook = new XSSFWorkbook(stream);

        CheckSheets(sheets, () => GetSheetNames(workbook));

        return sheets.Aggregate(new List<T>(), (current, sheetName) =>
        {
            var sheet = workbook.GetSheet(sheetName);
            var items = Read(sheet, options, mapperAction, itemAction, valueAction);
            current.AddRange(items);
            return current;
        });
    }

    protected virtual List<T> Read<T>(ISheet sheet, int firstRowNumber, int columnNumber, Func<object, CultureInfo, T> valueConvert = null)
    {
        var result = new List<T>();
        for (var index = firstRowNumber; index < sheet.LastRowNum; index++)
        {
            var row = sheet.GetRow(index);

            T value;

            if (valueConvert != null)
            {
                value = valueConvert(GetCellValue(row.GetCell(columnNumber, MissingCellPolicy.RETURN_NULL_AND_BLANK)), CultureInfo.CurrentCulture);
            }
            else
            {
                value = (T)GetCellValue(row.GetCell(columnNumber), typeof(T));
            }

            result.Add(value);
        }

        return result;
    }

    /// <summary>
    /// 读取表格内容并按行转换为对象集合
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="sheet"></param>
    /// <param name="options"></param>
    /// <param name="mapperAction"></param>
    /// <param name="itemAction"></param>
    /// <param name="valueAction"></param>
    /// <returns></returns>
    protected virtual List<T> Read<T>(ISheet sheet, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction)
    {
        if (sheet == null)
        {
            throw new NullReferenceException("The sheet can not be null.");
        }

        options ??= new SheetReadOptions();

        options.Validate();

        var mappers = new Dictionary<int, SheetColumnMapProfile>();

        var result = new List<T>();

        var headerRowNumber = GetHeaderRowNumber(sheet, options.HeaderRowNumber);

        var headerRow = sheet.GetRow(headerRowNumber);

        var columnCount = headerRow.Cells.Count;

        for (var index = options.FirstColumnNumber - 1; index < columnCount; index++)
        {
            var name = headerRow.GetCell(index)?.StringCellValue;
            if (string.IsNullOrWhiteSpace(name) || options.IgnoreNames.Contains(name, StringComparer.OrdinalIgnoreCase))
            {
                continue;
            }

            var mapper = options.GetMapProfile(name);
            mapper ??= new SheetColumnMapProfile(name, name);
            mappers.Add(index, mapper);
        }

        mapperAction?.Invoke(mappers);

        var firstDataRowNumber = headerRowNumber + 1;

        for (var rowIndex = firstDataRowNumber; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            if (result.Count >= options.RowCount)
            {
                break;
            }

            var row = sheet.GetRow(rowIndex);

            if (row == null)
            {
                continue;
            }

            var item = itemAction();

            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                if (!mappers.TryGetValue(columnIndex, out var mapper))
                {
                    continue;
                }

                var cellValue = GetCellValue(row.GetCell(columnIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK));
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
    /// 获取单元格值
    /// </summary>
    /// <param name="cell"></param>
    /// <param name="valueType"></param>
    /// <returns></returns>
    protected virtual object GetCellValue(ICell cell, Type valueType)
    {
        if (cell == null)
        {
            return null;
        }

        var value = GetCellValue(cell);

        if (value == null)
        {
            return null;
        }

        if (valueType == typeof(string))
        {
            return value.ToString();
        }

        if (valueType == typeof(int))
        {
            return int.Parse(value.ToString());
        }

        if (valueType == typeof(long))
        {
            return long.Parse(value.ToString());
        }

        if (valueType == typeof(DateTime))
        {
            return DateTime.Parse(value.ToString());
        }

        if (valueType == typeof(Guid))
        {
            return Guid.Parse(value.ToString());
        }

        return value;
    }

    /// <summary>
    /// 获取单元格值
    /// </summary>
    /// <param name="cell"></param>
    /// <returns></returns>
    protected virtual object GetCellValue(ICell cell)
    {
        if (cell == null)
        {
            return null;
        }

        return cell.CellType switch
        {
            CellType.String => cell.StringCellValue,
            CellType.Blank => string.Empty,
            CellType.Boolean => cell.BooleanCellValue,
            CellType.Error => cell.ErrorCellValue,
            CellType.Numeric => GetNumericCellValue(cell),
            CellType.Formula => string.Empty,
            CellType.Unknown => cell.StringCellValue,
            _ => string.Empty
        };
    }

    private static object GetNumericCellValue(ICell cell)
    {
        var stringValue = cell.ToString();
        if (stringValue.Contains("-") || stringValue.Contains(":"))
        {
            return cell.DateCellValue;
        }

        return cell.NumericCellValue;
    }

    private static IEnumerable<string> GetSheetNames(IWorkbook workbook)
    {
        for (var index = 0; index < workbook.NumberOfSheets; index++)
        {
            yield return workbook.GetSheetName(index);
        }
    }
}