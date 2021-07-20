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
    /// 表单内容解析契约接口
    /// </summary>
    public interface ISheetWrapper
    {
        /// <summary>
        /// 读取表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="options">读取配置</param>
        /// <param name="sheetIndex">表格位置索引</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="options">读取配置</param>
        /// <param name="sheetName">表格名称</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="options">读取配置</param>
        /// <param name="sheetIndex">表格位置索引</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="file"></param>
        /// <param name="options"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="file"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="options"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取单列数据为列表。
        /// </summary>
        /// <param name="file">文件路劲</param>
        /// <param name="firstRowNumber">（数据）首行行号</param>
        /// <param name="columnNumber">列号</param>
        /// <param name="sheetIndex">表格位置索引</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, int sheetIndex, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取单列数据为列表。
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="firstRowNumber">（数据）首行行号</param>
        /// <param name="columnNumber">列号</param>
        /// <param name="sheetIndex">表格位置索引</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, int sheetIndex, Func<object,  CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取单列数据到集合。
        /// </summary>
        /// <param name="file">文件路劲</param>
        /// <param name="firstRowNumber">（数据）首行行号</param>
        /// <param name="columnNumber">列号</param>
        /// <param name="sheetName">表格名称</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取单列数据为集合
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="firstRowNumber">（数据）首行行号</param>
        /// <param name="columnNumber">列号</param>
        /// <param name="sheetName">表格名称</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 将DataTable数据写入到表格并返回流。
        /// </summary>
        /// <param name="data"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<Stream> WriteAsync(DataTable data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default);

        /// <summary>
        /// 将对象写入到表格并返回流。
        /// </summary>
        /// <param name="data"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<Stream> WriteAsync<T>(IEnumerable<T> data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 将对象写入到表格并返回流
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataFactory"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<Stream> WriteAsync<T>(Func<Task<IEnumerable<T>>> dataFactory, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default
            where T : class, new();

        /// <summary>
        /// 将数据集合写入到表格的指定列并返回流。
        /// </summary>
        /// <param name="data"></param>
        /// <param name="firstRowNumber"></param>
        /// <param name="columnNumber"></param>
        /// <param name="valueConvert"></param>
        /// <param name="sheetName"></param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<Stream> WriteAsync<T>(IEnumerable<T> data, int firstRowNumber = 1, int columnNumber = 1, string sheetName = "", Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 将数据集合写入到表格的指定列并返回流
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataFactory"></param>
        /// <param name="firstRowNumber"></param>
        /// <param name="columnNumber"></param>
        /// <param name="sheetName"></param>
        /// <param name="valueConvert"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<Stream> WriteAsync<T>(Func<Task<IEnumerable<T>>> dataFactory, int firstRowNumber = 1, int columnNumber = 1, string sheetName = "", Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default);
    }
}