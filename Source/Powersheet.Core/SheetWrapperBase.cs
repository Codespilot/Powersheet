using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
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
        public abstract Task<Stream> WriteAsync<T>(IEnumerable<T> data, int firstRowNumber = 1, int columnNumber = 1, string sheetName = "", Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default);
    }
}