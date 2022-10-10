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
        public override async Task WriteAsync(Stream stream, DataTable data, SheetWriteOptions options, string sheetName, int itemsPerSheet, CancellationToken cancellationToken = default)
        {
            options ??= new SheetWriteOptions();
            options.Validate();

            await Task.Run(() =>
            {
                var mappers = GetColumnMapProfiles(data, options, true);

                var excel = new XSSFWorkbook();

                var style = excel.CreateCellStyle();
                var font = excel.CreateFont();

                var sheetCount = GetSheetCount(data.Rows.Count, ref itemsPerSheet);

                sheetName ??= "Sheet";

                for (var number = 0; number < sheetCount; number++)
                {
                    var name = $"{sheetName}{number + 1}";
                    var sheet = excel.CreateSheet(name);

                    var headerRowNumber = GetHeaderRowNumber(sheet, options.HeaderRowNumber);

                    var row = sheet.CreateRow(headerRowNumber);

                    foreach (var (index, profile) in mappers)
                    {
                        var cell = row.CreateCell(index, CellType.String);
                        cell.SetCellValue(profile.ColumnName);
                        cell.SetCellStyle(options.HeaderStyle, style, font);
                    }

                    var rows = data.Rows.OfType<DataRow>().Skip(number * itemsPerSheet).Take(itemsPerSheet);
                    Write(rows, sheet, options, mappers, GetValue, excel);
                }

                excel.Write(stream, true);
            }, cancellationToken);
        }

        /// <inherited/>
        public override async Task WriteAsync<T>(Stream stream, IEnumerable<T> data, SheetWriteOptions options, string sheetName, int itemsPerSheet, CancellationToken cancellationToken = default)
        {
            options ??= new SheetWriteOptions();
            options.Validate();
            SetStyle<T>(options);

            await Task.Run(() =>
            {
                var mappers = GetColumnMapProfiles<T>(options, true);

                var excel = new XSSFWorkbook();
                var style = excel.CreateCellStyle();
                var font = excel.CreateFont();

                var sheetCount = GetSheetCount(data.Count(), ref itemsPerSheet);

                sheetName ??= "Sheet";

                for (var number = 0; number < sheetCount; number++)
                {
                    var name = $"{sheetName}{number + 1}";
                    var sheet = excel.CreateSheet(name);

                    var headerRowNumber = GetHeaderRowNumber(sheet, options.HeaderRowNumber);

                    var row = sheet.CreateRow(headerRowNumber);

                    foreach (var (index, profile) in mappers)
                    {
                        var cell = row.CreateCell(index, CellType.String);
                        cell.SetCellValue(profile.ColumnName);
                        cell.SetCellStyle(options.HeaderStyle, style, font);
                    }

                    var rows = data.Skip(number * itemsPerSheet).Take(itemsPerSheet);

                    Write(rows, sheet, options, mappers, GetValue, excel);
                }

                excel.Write(stream, true);
            }, cancellationToken);
        }

        /// <inherited/>
        public override async Task<Stream> WriteAsync<T>(IEnumerable<T> data, int firstRowNumber, int columnNumber, string sheetName, Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default)
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

                    cell.SetCellValue(value);

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
        /// <param name="excel"></param>
        protected virtual void Write<T>(IEnumerable<T> data, ISheet sheet, SheetWriteOptions options, Dictionary<int, SheetColumnMapProfile> mappers, Func<T, string, object> valueAction, IWorkbook excel)
        {
            var style = excel.CreateCellStyle();
            var font = excel.CreateFont();

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
                        cellValue = mapper.ValueConverter.ConvertItemValue(sourceValue, CultureInfo.CurrentCulture);
                    }
                    else
                    {
                        cellValue = sourceValue;
                    }

                    var cell = row.CreateCell(columnNumber);
                    cell.SetCellValue(cellValue);
                    cell.SetCellStyle(options.BodyStyle, style, font);
                }

                currentRowNumber++;
            }
        }
    }
}