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
    public partial class SheetWrapper
    {
        /// <inherited/>
        public override async Task<Stream> WriteAsync(DataTable data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
        {
            options ??= new SheetWriteOptions();

            options.Validate();

            var columnsInDataTale = (from DataColumn column in data.Columns select column.ColumnName)
                                    .Where(name => !options.IgnoreNames.Contains(name, StringComparer.OrdinalIgnoreCase))
                                    .ToList();

            var mappers = new Dictionary<int, SheetColumnMapProfile>();

            var excel = new ExcelPackage();
            sheetName ??= "Sheet1";
            var sheet = excel.Workbook.Worksheets.Add(sheetName);
            for (var index = 0; index < columnsInDataTale.Count; index++)
            {
                var sheetColumnIndex = options.FirstColumnNumber + index;

                var name = columnsInDataTale[index];

                var mapper = options.GetMapProfile(name);
                mapper ??= new SheetColumnMapProfile(name, name);

                sheet.Cells[options.HeaderRowNumber, sheetColumnIndex].Value = mapper.ColumnName;

                mappers.Add(sheetColumnIndex, mapper);
            }

            static object GetValue(DataRow dataRow, string name)
            {
                return dataRow[name];
            }

            Write(data.Rows.OfType<DataRow>(), sheet, options, mappers, GetValue);

            var stream = new MemoryStream();
            await excel.SaveAsAsync(stream, cancellationToken);
            stream.Position = 0;
            stream.Seek(0, SeekOrigin.Begin);
            return stream;
        }

        /// <inherited/>
        public override async Task<Stream> WriteAsync<T>(IEnumerable<T> data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
        {
            options ??= new SheetWriteOptions();

            options.Validate();

            var properties = typeof(T).GetProperties();

            var mappers = new Dictionary<int, SheetColumnMapProfile>();

            var excel = new ExcelPackage();
            sheetName ??= "Sheet1";
            var sheet = excel.Workbook.Worksheets.Add(sheetName);

            var index = 0;

            foreach (var property in properties)
            {
                if (options.IgnoreNames.Contains(property.Name, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                var sheetColumnIndex = options.FirstColumnNumber + index;

                var name = property.Name;

                var mapper = options.GetMapProfile(name);
                mapper ??= new SheetColumnMapProfile(name, name);
                sheet.Cells[options.FirstColumnNumber, sheetColumnIndex].Value = mapper.ColumnName;
                mappers.Add(sheetColumnIndex, mapper);

                index++;
            }

            object GetValue(T item, string name)
            {
                return properties.FirstOrDefault(t => t.Name == name)?.GetValue(item);
            }

            Write(data, sheet, options, mappers, GetValue);

            var stream = new MemoryStream();
            await excel.SaveAsAsync(stream, cancellationToken);
            stream.Position = 0;
            stream.Seek(0, SeekOrigin.Begin);
            return stream;
        }

        /// <inherited/>
        public override async Task<Stream> WriteAsync<T>(IEnumerable<T> data, int firstRowNumber = 1, int columnNumber = 1, string sheetName = "", Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default)
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
                        cellValue = mapper.ValueConverter.ConvertBack(sourceValue, CultureInfo.CurrentCulture);
                    }
                    else
                    {
                        cellValue = sourceValue;
                    }

                    sheet.Cells[currentRowNumber, columnNumber].Value = cellValue;
                }

                currentRowNumber++;
            }
        }
    }
}