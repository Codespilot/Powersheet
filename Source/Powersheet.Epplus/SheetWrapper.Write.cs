using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Nerosoft.Powersheet.Epplus
{
    /// <summary>
    /// 基于EPPLUS实现表格解析
    /// </summary>
    public partial class SheetWrapper
    {
        /// <inherited/>
        public override async Task WriteAsync(Stream stream, DataTable data, SheetWriteOptions options, string sheetName, int itemsPerSheet, CancellationToken cancellationToken = default)
        {
            options ??= new SheetWriteOptions();

            options.Validate();

            var mappers = GetColumnMapProfiles(data, options, false);

            var excel = new ExcelPackage();

            var sheetCount = GetSheetCount(data.Rows.Count, ref itemsPerSheet);

            sheetName ??= "Sheet";
            for (var number = 0; number < sheetCount; number++)
            {
                var name = $"{sheetName}{number + 1}";
                var sheet = excel.Workbook.Worksheets.Add(name);

                foreach (var (index, profile) in mappers)
                {
                    var cell = sheet.Cells[options.FirstColumnNumber, index];
                    cell.Value = profile.ColumnName;
                    cell.SetCellStyle(options.HeaderStyle);
                }

                var rows = data.Rows.OfType<DataRow>().Skip(number * itemsPerSheet).Take(itemsPerSheet);
                Write(rows, sheet, options, mappers, GetValue);
            }

            await excel.SaveAsAsync(stream, cancellationToken);
        }

        /// <inherited/>
        public override async Task WriteAsync<T>(Stream stream, IEnumerable<T> data, SheetWriteOptions options, string sheetName, int itemsPerSheet, CancellationToken cancellationToken = default)
        {
            options ??= new SheetWriteOptions();
            options.Validate();
            SetStyle<T>(options);

            var mappers = GetColumnMapProfiles<T>(options, false);

            var excel = new ExcelPackage();

            var sheetCount = GetSheetCount(data.Count(), ref itemsPerSheet);

            sheetName ??= "Sheet";

            for (var number = 0; number < sheetCount; number++)
            {
                var name = $"{sheetName}{number + 1}";
                var sheet = excel.Workbook.Worksheets.Add(name);

                foreach (var (index, profile) in mappers)
                {
                    var cell = sheet.Cells[options.FirstColumnNumber, index];
                    cell.Value = profile.ColumnName;
                    cell.SetCellStyle(options.HeaderStyle);
                }

                var rows = data.Skip(number * itemsPerSheet).Take(itemsPerSheet);
                Write(rows, sheet, options, mappers, GetValue);
            }

            await excel.SaveAsAsync(stream, cancellationToken);
        }

        /// <inherited/>
        public override async Task<Stream> WriteAsync<T>(IEnumerable<T> data, int firstRowNumber, int columnNumber, string sheetName, Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default)
        {
            var excel = new ExcelPackage();
            sheetName ??= "Sheet1";
            var sheet = excel.Workbook.Worksheets.Add(sheetName);

            var currentRowNumber = firstRowNumber;

            foreach (var item in data)
            {
                var value = valueConvert == null ? item : valueConvert(item, CultureInfo.CurrentCulture);
                sheet.Cells[currentRowNumber, columnNumber].Value = value;
                currentRowNumber++;
            }

            var stream = new MemoryStream();
            await excel.SaveAsAsync(stream, cancellationToken);
            stream.Position = 0;
            stream.Seek(0, SeekOrigin.Begin);
            return stream;
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
        protected virtual void Write<T>(IEnumerable<T> data, ExcelWorksheet sheet, SheetWriteOptions options, Dictionary<int, SheetColumnMapProfile> mappers, Func<T, string, object> valueAction)
        {
            var currentRowNumber = options.HeaderRowNumber + 1;
            foreach (var item in data)
            {
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

                    var cell = sheet.Cells[currentRowNumber, columnNumber];
                    cell.Value = cellValue;
                    cell.SetCellStyle(options.BodyStyle);
                }

                currentRowNumber++;
            }
        }
    }
}