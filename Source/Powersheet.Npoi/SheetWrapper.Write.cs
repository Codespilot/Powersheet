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
        /// <inherited/>
        public override async Task<Stream> WriteAsync(DataTable data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
        {
            options ??= new SheetWriteOptions();
            options.Validate();

            return await Task.Run(() =>
            {
                var columnsInDataTale = (from DataColumn column in data.Columns select column.ColumnName)
                                        .Where(name => !options.IgnoreNames.Contains(name, StringComparer.OrdinalIgnoreCase))
                                        .ToList();

                var mappers = new Dictionary<int, SheetColumnMapProfile>();

                var excel = new XSSFWorkbook();
                var sheet = excel.CreateSheet(sheetName);

                var headerRowNumber = GetHeaderRowNumber(sheet, options.HeaderRowNumber);

                var row = sheet.CreateRow(headerRowNumber);

                for (var index = 0; index < columnsInDataTale.Count; index++)
                {
                    var sheetColumnIndex = options.FirstColumnNumber + index - 1;

                    var name = columnsInDataTale[index];

                    var mapper = options.GetMapProfile(name);
                    mapper ??= new SheetColumnMapProfile(name, name);

                    row.CreateCell(sheetColumnIndex, CellType.String).SetCellValue(mapper.ColumnName);

                    mappers.Add(sheetColumnIndex, mapper);
                }

                static object GetValue(DataRow dataRow, string name)
                {
                    return dataRow[name];
                }

                Write(data.Rows.OfType<DataRow>(), sheet, options, mappers, GetValue);

                var stream = new MemoryStream();
                excel.Write(stream, true);
                stream.Position = 0;
                stream.Seek(0, SeekOrigin.Begin);
                return stream;
            }, cancellationToken);
        }

        /// <inherited/>
        public override async Task<Stream> WriteAsync<T>(IEnumerable<T> data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
        {
            options ??= new SheetWriteOptions();
            options.Validate();

            return await Task.Run(() =>
            {
                var properties = typeof(T).GetProperties();

                var mappers = new Dictionary<int, SheetColumnMapProfile>();

                var excel = new XSSFWorkbook();
                var sheet = excel.CreateSheet(sheetName);

                var headerRowNumber = GetHeaderRowNumber(sheet, options.HeaderRowNumber);

                var row = sheet.CreateRow(headerRowNumber);

                var index = 0;

                foreach (var property in properties)
                {
                    if (options.IgnoreNames.Contains(property.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var sheetColumnIndex = options.FirstColumnNumber + index - 1;

                    var name = property.Name;
                    var mapper = options.GetMapProfile(name);
                    mapper ??= new SheetColumnMapProfile(name, name);
                    row.CreateCell(sheetColumnIndex, CellType.String).SetCellValue(mapper.ColumnName);
                    mappers.Add(sheetColumnIndex, mapper);

                    index++;
                }

                object GetValue(T item, string name)
                {
                    return properties.FirstOrDefault(t => t.Name == name)?.GetValue(item);
                }

                Write(data, sheet, options, mappers, GetValue);

                var stream = new MemoryStream();
                excel.Write(stream, true);
                stream.Position = 0;
                stream.Seek(0, SeekOrigin.Begin);
                return stream;
            }, cancellationToken);
        }

        /// <inherited/>
        public override async Task<Stream> WriteAsync<T>(IEnumerable<T> data, int firstRowNumber = 1, int columnNumber = 1, string sheetName = "", Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() =>
            {
                var excel = new XSSFWorkbook();
                var sheet = excel.CreateSheet(sheetName);

                var currentRowNumber = firstRowNumber - 1;
                foreach (var item in data)
                {
                    var row = sheet.CreateRow(currentRowNumber);
                    var value = valueConvert != null ? valueConvert(item, CultureInfo.CurrentCulture) : item;

                    var cell = row.CreateCell(columnNumber - 1);

                    SetCellValue(cell, value);

                    currentRowNumber++;
                }

                var stream = new MemoryStream();
                excel.Write(stream, true);
                stream.Position = 0;
                stream.Seek(0, SeekOrigin.Begin);
                return stream;
            }, cancellationToken);
        }

        /// <summary>
        /// 将数据集合写入到表格
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="sheet"></param>
        /// <param name="options"></param>
        /// <param name="mappers"></param>
        /// <param name="valueAction"></param>
        protected virtual void Write<T>(IEnumerable<T> data, ISheet sheet, SheetWriteOptions options, Dictionary<int, SheetColumnMapProfile> mappers, Func<T, string, object> valueAction)
        {
            var currentRowNumber = GetHeaderRowNumber(sheet, options.HeaderRowNumber) + 1;

            foreach (var item in data)
            {
                var row = sheet.CreateRow(currentRowNumber);
                foreach (var (columnNumber, mapper) in mappers)
                {
                    var sourceValue = valueAction.Invoke(item, mapper.Name);
                    object cellValue;
                    if (mapper.ValueConvert != null)
                    {
                        cellValue = mapper.ValueConvert(sourceValue, CultureInfo.CurrentCulture);
                    }
                    else if (mapper.ValueConverter != null)
                    {
                        cellValue = mapper.ValueConverter.ConvertBack(sourceValue, CultureInfo.CurrentCulture);
                    }
                    else
                    {
                        cellValue = sourceValue;
                    }

                    var cell = row.CreateCell(columnNumber);

                    SetCellValue(cell, cellValue);
                }

                currentRowNumber++;
            }
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        protected virtual void SetCellValue(ICell cell, object value)
        {
            switch (value)
            {
                case null:
                    return;
                case string cellValue:
                    cell.SetCellValue(cellValue);
                    break;
                case int:
                case long:
                case decimal:
                case double:
                    cell.SetCellValue(value.ToString());
                    cell.SetCellType(CellType.Numeric);
                    break;
                case DateTime cellValue:
                    cell.SetCellValue(cellValue);
                    break;
                case bool cellValue:
                    cell.SetCellValue(cellValue);
                    break;
                default:
                    cell.SetCellValue(value.ToString());
                    break;
            }
        }
    }
}