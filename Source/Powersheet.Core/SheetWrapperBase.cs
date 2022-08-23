using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace Nerosoft.Powersheet
{
    /// <summary>
    /// 表格解析基类
    /// </summary>
    public abstract class SheetWrapperBase : ISheetWrapper
    {
        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
        {
            return await ReadToDataTableAsync(file, options, new[] { sheetIndex }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
        {
            return await ReadToDataTableAsync(file, options, new[] { sheetName }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, IEnumerable<int> sheets, CancellationToken cancellationToken = default)
        {
            using (var stream = File.OpenRead(file))
            {
                return await ReadToDataTableAsync(stream, options, sheets, cancellationToken);
            }
        }

        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, IEnumerable<string> sheets, CancellationToken cancellationToken = default)
        {
            using (var stream = File.OpenRead(file))
            {
                return await ReadToDataTableAsync(stream, options, sheets, cancellationToken);
            }
        }

        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
        {
            return await ReadToDataTableAsync(stream, options, new[] { sheetIndex }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
        {
            return await ReadToDataTableAsync(stream, options, new[] { sheetName }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, IEnumerable<int> sheets, CancellationToken cancellationToken = default)
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

                var rows = Read(stream, options, MapperAction, () => table.NewRow(), ValueAction, sheets);

                foreach (var row in rows)
                {
                    table.Rows.Add(row);
                }

                return table;
            }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, IEnumerable<string> sheets, CancellationToken cancellationToken = default)
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

                var rows = Read(stream, options, MapperAction, () => table.NewRow(), ValueAction, sheets);

                foreach (var row in rows)
                {
                    table.Rows.Add(row);
                }

                return table;
            }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            return await ReadToListAsync<T>(file, options, new[] { sheetIndex }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, IEnumerable<int> sheets, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            using (var stream = File.OpenRead(file))
            {
                return await ReadToListAsync<T>(stream, options, sheets, cancellationToken);
            }
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            return await ReadToListAsync<T>(file, options, new[] { sheetName }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, IEnumerable<string> sheets, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            using (var stream = File.OpenRead(file))
            {
                return await ReadToListAsync<T>(stream, options, sheets, cancellationToken);
            }
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            return await ReadToListAsync<T>(stream, options, new[] { sheetIndex }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, IEnumerable<int> sheets, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            options ??= new SheetReadOptions();

            options.Validate();

            MergeMapProfile<T>(options);

            var properties = typeof(T).GetProperties().ToDictionary(t => t.Name, t => t);

            void SetItemValue(T item, string name, object value) => SetValue(properties, item, name, value);

            return await Task.Run(() =>
            {
                var items = Read(stream, options, null, () => new T(), SetItemValue, sheets);
                return items;
            }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            return await ReadToListAsync<T>(stream, options, new[] { sheetName }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, IEnumerable<string> sheets, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            options ??= new SheetReadOptions();

            options.Validate();

            MergeMapProfile<T>(options);

            var properties = typeof(T).GetProperties().ToDictionary(t => t.Name, t => t);

            void SetItemValue(T item, string name, object value) => SetValue(properties, item, name, value);

            return await Task.Run(() =>
            {
                var items = Read(stream, options, null, () => new T(), SetItemValue, sheets);
                return items;
            }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, int sheetIndex, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
        {
            return await ReadToListAsync(file, firstRowNumber, columnNumber, new[] { sheetIndex }, valueConvert, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
        {
            return await ReadToListAsync(file, firstRowNumber, columnNumber, new[] { sheetName }, valueConvert, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, IEnumerable<int> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
        {
            using (var stream = File.OpenRead(file))
            {
                return await ReadToListAsync(stream, firstRowNumber, columnNumber, sheets, valueConvert, cancellationToken);
            }
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, IEnumerable<string> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
        {
            using (var stream = File.OpenRead(file))
            {
                return await ReadToListAsync(stream, firstRowNumber, columnNumber, sheets, valueConvert, cancellationToken);
            }
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, int sheetIndex, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
        {
            return await ReadToListAsync(stream, firstRowNumber, columnNumber, new[] { sheetIndex }, valueConvert, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
        {
            return await ReadToListAsync(stream, firstRowNumber, columnNumber, new[] { sheetName }, valueConvert, cancellationToken);
        }

        /// <inherited/>
        public abstract Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, IEnumerable<int> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <inherited/>
        public abstract Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, IEnumerable<string> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        protected abstract List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, IEnumerable<int> sheets);

        protected abstract List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, IEnumerable<string> sheets);

        public abstract Task WriteAsync(Stream stream, DataTable data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default);

        public abstract Task WriteAsync<T>(Stream stream, IEnumerable<T> data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();

        public virtual async Task WriteAsync<T>(Stream stream, Func<Task<IEnumerable<T>>> dataFactory, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            if (dataFactory == null)
            {
                throw new ArgumentNullException(nameof(dataFactory));
            }

            var data = await dataFactory();
            await WriteAsync(stream, data, options, sheetName, cancellationToken);
        }

        /// <inherited/>
        public abstract Task<Stream> WriteAsync(DataTable data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default);

        /// <inherited/>
        public abstract Task<Stream> WriteAsync<T>(IEnumerable<T> data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default) where T : class, new();

        /// <inherited/>
        public virtual async Task<Stream> WriteAsync<T>(Func<Task<IEnumerable<T>>> dataFactory, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            if (dataFactory == null)
            {
                throw new ArgumentNullException(nameof(dataFactory));
            }

            var data = await dataFactory();
            return await WriteAsync(data, options, sheetName, cancellationToken);
        }

        /// <inherited/>
        public abstract Task<Stream> WriteAsync<T>(IEnumerable<T> data, int firstRowNumber, int columnNumber, string sheetName, Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default);

        /// <inherited/>
        public virtual async Task<Stream> WriteAsync<T>(Func<Task<IEnumerable<T>>> dataFactory, int firstRowNumber, int columnNumber, string sheetName, Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default)
        {
            if (dataFactory == null)
            {
                throw new ArgumentNullException(nameof(dataFactory));
            }

            var data = await dataFactory();
            return await WriteAsync(data, firstRowNumber, columnNumber, sheetName, valueConvert, cancellationToken);
        }

        /// <summary>
        /// 合并对象属性映射配置
        /// </summary>
        /// <param name="options"></param>
        /// <typeparam name="T"></typeparam>
        protected virtual void MergeMapProfile<T>(SheetHandleOptions options)
            where T : class
        {
            var properties = typeof(T).GetProperties();

            var attributes = properties.ToDictionary(t => t.Name, t => t.GetCustomAttribute<SheetColumnMapAttribute>());

            var duplicates = attributes.Where(t => t.Value != null)
                                       .GroupBy(t => t.Value.ColumnName)
                                       .Where(t => t.Count() > 1)
                                       .Select(t => t.Key);

            if (duplicates.Any())
            {
                throw new DuplicateNameException($"Duplicate column name in attribute map profile: {string.Join(",", duplicates)}");
            }

            foreach (var property in properties)
            {
                if (options.Mapping.Any(t => t.Name == property.Name))
                {
                    continue;
                }

                if (!attributes.TryGetValue(property.Name, out var attribute))
                {
                    continue;
                }

                if (attribute == null)
                {
                    continue;
                }

                if (options.Mapping.Any(t => t.ColumnName == attribute.ColumnName))
                {
                    continue;
                }

                var profile = new SheetColumnMapProfile(property.Name, attribute.ColumnName, attribute.ValueConverter);

                options.AddMapProfile(profile);
            }
        }

        /// <summary>
        /// 设置样式
        /// </summary>
        /// <param name="options"></param>
        /// <typeparam name="T"></typeparam>
        protected virtual void SetStyle<T>(SheetWriteOptions options)
            where T : class
        {
            var type = typeof(T);

            if (options.HeaderStyle == null)
            {
                var attribute = type.GetCustomAttribute<SheetHeaderStyleAttribute>();
                if (attribute != null)
                {
                    options.HeaderStyle = attribute.Style;
                }
                else
                {
                    var attributes = type.GetCustomAttributes().OfType<ISheetStyleAttribute>();

                    if (attributes.Any())
                    {
                        options.HeaderStyle = new CellStyle();
                    }

                    foreach (var attr in attributes)
                    {
                        attr.SetStyle(options.HeaderStyle);
                    }
                }
            }

            if (options.BodyStyle == null)
            {
                var attribute = type.GetCustomAttribute<SheetBodyStyleAttribute>();
                if (attribute != null)
                {
                    options.BodyStyle = attribute.Style;
                }
                else
                {
                    var attributes = type.GetCustomAttributes().OfType<ISheetStyleAttribute>();

                    if (attributes.Any())
                    {
                        options.BodyStyle = new CellStyle();
                    }

                    foreach (var attr in attributes)
                    {
                        attr.SetStyle(options.BodyStyle);
                    }
                }
            }
        }

        /// <summary>
        /// 设置对象属性值
        /// </summary>
        /// <param name="properties">属性字典</param>
        /// <param name="item">设置的对象</param>
        /// <param name="name">属性名称</param>
        /// <param name="value">属性值</param>
        /// <typeparam name="T">对象类型</typeparam>
        protected virtual void SetValue<T>(IDictionary<string, PropertyInfo> properties, T item, string name, object value)
        {
            if (!properties.TryGetValue(name, out var property))
            {
                return;
            }

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

        protected virtual void CheckSheets(IEnumerable<int> sheets)
        {
            if (sheets == null)
            {
                throw new ArgumentNullException(nameof(sheets));
            }

            if (!sheets.Any())
            {
                throw new ArgumentException("Sheets is empty.", nameof(sheets));
            }

            if (sheets.Any(t => t < 0))
            {
                throw new IndexOutOfRangeException("The sheet index could not less than 0.");
            }
        }

        protected virtual void CheckSheets(IEnumerable<int> sheets, int sheetCount)
        {
            if (sheets.Any(sheetIndex => sheetIndex >= sheetCount))
            {
                throw new IndexOutOfRangeException($"The sheet index could not greater than or equals the worksheet count ({sheetCount}).");
            }
        }

        protected virtual void CheckSheets(IEnumerable<string> sheets)
        {
            if (sheets == null)
            {
                throw new ArgumentNullException(nameof(sheets));
            }

            if (!sheets.Any())
            {
                throw new ArgumentException("Sheets is empty.", nameof(sheets));
            }

            if (sheets.Any(string.IsNullOrEmpty))
            {
                throw new InvalidOperationException("The sheet name could not be null or empty.");
            }
        }

        protected virtual void CheckSheets(IEnumerable<string> sheets, Func<IEnumerable<string>> getSheetNames)
        {
            var names = getSheetNames();
            var notContainsSheets = sheets.Where(t => !names.Contains(t));
            if (notContainsSheets.Any())
            {
                throw new SheetNotFoundException(string.Join(",", notContainsSheets), $"The workbook does not contains the sheet(s) named '{string.Join(",", notContainsSheets)}'");
            }
        }
    }
}