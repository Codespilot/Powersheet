using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Nerosoft.Powersheet.Npoi
{
    public partial class SheetWrapper
    {
        public override async Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() =>
            {
                var table = new DataTable();

                void MapperAction(Dictionary<int, SheetColumnMapProfile> mapper)
                {
                    foreach (var item in mapper.Values)
                    {
                        table.Columns.Add(item.Name);
                    }
                }

                static void ValueAction(DataRow item, string name, object value)
                {
                    item[name] = value ?? DBNull.Value;
                }

                var rows = Read(stream, options, MapperAction, () => table.NewRow(), ValueAction, sheetIndex);

                foreach (var row in rows)
                {
                    table.Rows.Add(row);
                }

                return table;
            }, cancellationToken);
        }

        public override async Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() =>
            {
                var table = new DataTable();

                void MapperAction(Dictionary<int, SheetColumnMapProfile> mapper)
                {
                    foreach (var item in mapper.Values)
                    {
                        table.Columns.Add(item.Name);
                    }
                }

                void ValueAction(DataRow item, string name, object value)
                {
                    item[name] = value ?? DBNull.Value;
                }

                var rows = Read(stream, options, MapperAction, () => table.NewRow(), ValueAction, sheetName);

                foreach (var row in rows)
                {
                    table.Rows.Add(row);
                }

                return table;
            }, cancellationToken);
        }

        public override async Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
        {
            var properties = typeof(T).GetProperties().ToDictionary(t => t.Name, t => t);

            void ValueAction(T item, string name, object value)
            {
                var property = properties[name];
                if (property == null || !property.CanWrite || value == null)
                {
                    return;
                }

                if (value.GetType() != property.PropertyType)
                {
                    value = Convert.ChangeType(value, property.PropertyType);
                }

                property.SetValue(item, value);
            }

            return await Task.Run(() =>
            {
                var items = Read(stream, options, null, () => new T(), ValueAction, sheetIndex);
                return items;
            }, cancellationToken);
        }

        public override async Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
        {
            var properties = typeof(T).GetProperties().ToDictionary(t => t.Name, t => t);

            void ValueAction(T item, string name, object value)
            {
                var property = properties[name];
                if (property == null || !property.CanWrite || value == null)
                {
                    return;
                }

                if (value.GetType() != property.PropertyType)
                {
                    value = Convert.ChangeType(value, property.PropertyType);
                }

                property.SetValue(item, value);
            }

            return await Task.Run(() =>
            {
                var items = Read(stream, options, null, () => new T(), ValueAction, sheetName);
                return items;
            }, cancellationToken);
        }

        public override async Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, int sheetIndex, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() =>
            {
                var excel = new XSSFWorkbook(stream);
                var sheet = excel.GetSheetAt(sheetIndex);

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
                        value = (T) GetCellValue(row.GetCell(columnNumber), typeof(T));
                    }

                    result.Add(value);
                }

                return result;
            }, cancellationToken);
        }

        public override async Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() =>
            {
                var excel = new XSSFWorkbook(stream);
                var sheet = excel.GetSheet(sheetName);

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
                        value = (T) GetCellValue(row.GetCell(columnNumber), typeof(T));
                    }

                    result.Add(value);
                }

                return result;
            }, cancellationToken);
        }

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
                if (string.IsNullOrWhiteSpace(name) || options.IgnoreColumns.Contains(name, StringComparer.OrdinalIgnoreCase))
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

                if (row.Cells.All(t => string.IsNullOrWhiteSpace(GetCellValue(t, typeof(string)) as string)))
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
                        value = mapper.ValueConverter.Convert(cellValue, CultureInfo.CurrentCulture);
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

        protected virtual List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, int sheetIndex = 0)
        {
            if (sheetIndex < 0)
            {
                throw new IndexOutOfRangeException("The sheet index could not less than 0.");
            }

            var excel = new XSSFWorkbook(stream);
            if (sheetIndex >= excel.NumberOfSheets)
            {
                throw new IndexOutOfRangeException($"The sheet index could not larger than or equals the workbook sheet count ({excel.NumberOfNames}).");
            }

            var sheet = excel.GetSheetAt(sheetIndex);

            return Read(sheet, options, mapperAction, itemAction, valueAction);
        }

        protected virtual List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }

            var excel = new XSSFWorkbook(stream);
            ISheet sheet;
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                var names = GetSheetNames();
                if (names.All(t => t.Equals(sheetName)))
                {
                    throw new InvalidOperationException($"The workbook does not contains a sheet named '{sheetName}'");
                }

                sheet = excel.GetSheet(sheetName);
            }
            else
            {
                sheet = excel.GetSheetAt(0);
            }

            return Read(sheet, options, mapperAction, itemAction, valueAction);

            IEnumerable<string> GetSheetNames()
            {
                for (var index = 0; index < excel.NumberOfSheets; index++)
                {
                    yield return excel.GetSheetName(index);
                }
            }
        }

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

        
    }
}