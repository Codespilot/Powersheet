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
        public abstract Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default);

        /// <inherited/>
        public abstract Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default);

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
        public abstract Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <inherited/>
        public abstract Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();

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