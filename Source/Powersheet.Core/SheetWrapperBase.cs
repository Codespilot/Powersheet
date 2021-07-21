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
            var stream = File.OpenRead(file);
            return await ReadToDataTableAsync(stream, options, sheetIndex, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
        {
            var stream = File.OpenRead(file);
            return await ReadToDataTableAsync(stream, options, sheetName, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
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

        /// <inherited/>
        public virtual async Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
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

                var rows = Read(stream, options, MapperAction, () => table.NewRow(), ValueAction, sheetName);

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
            var stream = File.OpenRead(file);
            return await ReadToListAsync<T>(stream, options, sheetIndex, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            var stream = File.OpenRead(file);
            return await ReadToListAsync<T>(stream, options, sheetName, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            options ??= new SheetReadOptions();

            options.Validate();

            MergeMapProfile<T>(options);

            var properties = typeof(T).GetProperties().ToDictionary(t => t.Name, t => t);

            void SetItemValue(T item, string name, object value) => SetValue(properties, item, name, value);

            return await Task.Run(() =>
            {
                var items = Read(stream, options, null, () => new T(), SetItemValue, sheetIndex);
                return items;
            }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            options ??= new SheetReadOptions();

            options.Validate();

            MergeMapProfile<T>(options);

            var properties = typeof(T).GetProperties().ToDictionary(t => t.Name, t => t);

            void SetItemValue(T item, string name, object value) => SetValue(properties, item, name, value);

            return await Task.Run(() =>
            {
                var items = Read(stream, options, null, () => new T(), SetItemValue, sheetName);
                return items;
            }, cancellationToken);
        }

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, int sheetIndex, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
        {
            var stream = File.OpenRead(file);
            return await ReadToListAsync(stream, firstRowNumber, columnNumber, sheetIndex, valueConvert, cancellationToken);
        }

        /// <inherited/>
        public abstract Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, int sheetIndex, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <inherited/>
        public virtual async Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default)
        {
            var stream = File.OpenRead(file);
            return await ReadToListAsync(stream, firstRowNumber, columnNumber, sheetName, valueConvert, cancellationToken);
        }

        /// <inherited/>
        public abstract Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取表格内容并按行转换为对象集合
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="options"></param>
        /// <param name="mapperAction"></param>
        /// <param name="itemAction"></param>
        /// <param name="valueAction"></param>
        /// <param name="sheetIndex"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        protected abstract List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, int sheetIndex = 0);

        /// <summary>
        /// 读取表格内容并按行转换为对象集合
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="options"></param>
        /// <param name="mapperAction"></param>
        /// <param name="itemAction"></param>
        /// <param name="valueAction"></param>
        /// <param name="sheetName"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        protected abstract List<T> Read<T>(Stream stream, SheetReadOptions options, Action<Dictionary<int, SheetColumnMapProfile>> mapperAction, Func<T> itemAction, Action<T, string, object> valueAction, string sheetName);
        
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
        public abstract Task<Stream> WriteAsync<T>(IEnumerable<T> data, int firstRowNumber = 1, int columnNumber = 1, string sheetName = "", Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default);

        /// <inherited/>
        public virtual async Task<Stream> WriteAsync<T>(Func<Task<IEnumerable<T>>> dataFactory, int firstRowNumber = 1, int columnNumber = 1, string sheetName = "", Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default)
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
                    return;
                }

                if (attribute == null)
                {
                    return;
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
        /// 设置对象属性值
        /// </summary>
        /// <param name="properties"></param>
        /// <param name="item"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <typeparam name="T"></typeparam>
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
    }
}