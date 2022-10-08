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
        #region ReadToDataTable

        /// <summary>
        /// 读取表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetIndex">表格位置索引，起始值为0</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetIndex">表格位置索引，起始值为0</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取多个表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheets">表格位置索引，起始值为0</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, IEnumerable<int> sheets, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheets">表格名称，不指定取第一个</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(string file, SheetReadOptions options, IEnumerable<string> sheets, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取多个表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="stream">文件路径</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheets">表格位置索引，起始值为0</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, IEnumerable<int> sheets, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取表格内容到<see cref="DataTable"/>。
        /// </summary>
        /// <param name="stream">文件路径</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheets">表格名称，不指定取第一个</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<DataTable> ReadToDataTableAsync(Stream stream, SheetReadOptions options, IEnumerable<string> sheets, CancellationToken cancellationToken = default);

        #endregion

        #region ReadToList

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetIndex">表格位置索引，起始值为0</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型，必须是类且包含公开的无参构造器</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="options"></param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型，必须是类且包含公开的无参构造器</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheets">表格位置索引，起始值为0</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型，必须是类且包含公开的无参构造器</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, IEnumerable<int> sheets, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="options"></param>
        /// <param name="sheets">表格名称</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型，必须是类且包含公开的无参构造器</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, SheetReadOptions options, IEnumerable<string> sheets, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetIndex">表格位置索引，起始值为0</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型，必须是类且包含公开的无参构造器</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, int sheetIndex, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型，必须是类且包含公开的无参构造器</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheets">表格位置索引，起始值为0</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型，必须是类且包含公开的无参构造器</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, IEnumerable<int> sheets, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取表格内容到对象集合
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheets">表格名称</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型，必须是类且包含公开的无参构造器</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, SheetReadOptions options, IEnumerable<string> sheets, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 读取指定单列数据到集合
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="firstRowNumber">起始行号，从1开始</param>
        /// <param name="columnNumber">列号，起始值为1</param>
        /// <param name="sheetIndex">表格位置索引，起始值为0</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, int sheetIndex, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取指定单列数据到集合
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="firstRowNumber">起始行号，从1开始</param>
        /// <param name="columnNumber">列号，起始值为1</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取指定单列数据到集合
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="firstRowNumber">起始行号，从1开始</param>
        /// <param name="columnNumber">列号，起始值为1</param>
        /// <param name="sheets">表格位置索引，起始值为0</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, IEnumerable<int> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取指定单列数据到集合
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="firstRowNumber">起始行号，从1开始</param>
        /// <param name="columnNumber">列号，起始值为1</param>
        /// <param name="sheets">表格名称</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(string file, int firstRowNumber, int columnNumber, IEnumerable<string> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取指定单列数据到集合
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="firstRowNumber">起始行号，从1开始</param>
        /// <param name="columnNumber">列号，起始值为1</param>
        /// <param name="sheetIndex">表格位置索引，起始值为0</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, int sheetIndex, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取指定单列数据到集合
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="firstRowNumber">起始行号，从1开始</param>
        /// <param name="columnNumber">列号，起始值为1</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, string sheetName, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取指定单列数据到集合
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="firstRowNumber">起始行号，从1开始</param>
        /// <param name="columnNumber">列号，起始值为1</param>
        /// <param name="sheets">表格位置索引</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, IEnumerable<int> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 读取指定单列数据到集合
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="firstRowNumber">起始行号，从1开始</param>
        /// <param name="columnNumber">列号，起始值为1</param>
        /// <param name="sheets">表格名称</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">结果对象类型</typeparam>
        /// <returns></returns>
        Task<List<T>> ReadToListAsync<T>(Stream stream, int firstRowNumber, int columnNumber, IEnumerable<string> sheets, Func<object, CultureInfo, T> valueConvert = null, CancellationToken cancellationToken = default);

        #endregion

        #region Write

        /// <summary>
        /// 将DataTable数据写入到流
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="data"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task WriteAsync(Stream stream, DataTable data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default);

        /// <summary>
        /// 将DataTable数据写入到流
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="data"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="itemsPerSheet"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task WriteAsync(Stream stream, DataTable data, SheetWriteOptions options, string sheetName, int itemsPerSheet, CancellationToken cancellationToken = default);

        /// <summary>
        /// 将对象集合转换为表格后写入到流
        /// </summary>
        /// <param name="stream">目标流</param>
        /// <param name="data">要写入到流的数据集合</param>
        /// <param name="options">写入选项</param>
        /// <param name="sheetName">表格名称</param>
        /// <param name="cancellationToken">操作取消令牌</param>
        /// <typeparam name="T">对象类型</typeparam>
        /// <returns></returns>
        Task WriteAsync<T>(Stream stream, IEnumerable<T> data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 将对象集合转换为表格后写入到流
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="data"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="itemsPerSheet"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task WriteAsync<T>(Stream stream, IEnumerable<T> data, SheetWriteOptions options, string sheetName, int itemsPerSheet, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 将对象集合转换为表格后写入到流
        /// </summary>
        /// <param name="stream">目标流</param>
        /// <param name="dataFactory">取得要写入到流的数据集合的方法委托</param>
        /// <param name="options">写入选项</param>
        /// <param name="sheetName">表格名称</param>
        /// <param name="cancellationToken">操作取消令牌</param>
        /// <typeparam name="T">对象类型</typeparam>
        /// <returns></returns>
        Task WriteAsync<T>(Stream stream, Func<Task<IEnumerable<T>>> dataFactory, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 将DataTable数据写入到表格并返回流。
        /// </summary>
        /// <param name="data">数据集合</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<Stream> WriteAsync(DataTable data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default);
        
        /// <summary>
        /// 将DataTable数据写入到表格并返回流。
        /// </summary>
        /// <param name="data">数据集合</param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="itemsPerSheet">每个表格写入的行数</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        Task<Stream> WriteAsync(DataTable data, SheetWriteOptions options, string sheetName, int itemsPerSheet, CancellationToken cancellationToken = default);

        /// <summary>
        /// 将对象写入到表格并返回流。
        /// </summary>
        /// <param name="data"></param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">数据对象类型</typeparam>
        /// <returns></returns>
        Task<Stream> WriteAsync<T>(IEnumerable<T> data, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();
        
        /// <summary>
        /// 将对象写入到表格并返回流。
        /// </summary>
        /// <param name="data"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="itemsPerSheet"></param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<Stream> WriteAsync<T>(IEnumerable<T> data, SheetWriteOptions options, string sheetName, int itemsPerSheet, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 将对象写入到表格并返回流
        /// </summary>
        /// <param name="dataFactory"></param>
        /// <param name="options">配置选项</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">数据对象类型</typeparam>
        /// <returns></returns>
        Task<Stream> WriteAsync<T>(Func<Task<IEnumerable<T>>> dataFactory, SheetWriteOptions options, string sheetName, CancellationToken cancellationToken = default)
            where T : class, new();
        
        /// <summary>
        /// 将对象写入到表格并返回流
        /// </summary>
        /// <param name="dataFactory"></param>
        /// <param name="options"></param>
        /// <param name="sheetName"></param>
        /// <param name="itemsPerSheet"></param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<Stream> WriteAsync<T>(Func<Task<IEnumerable<T>>> dataFactory, SheetWriteOptions options, string sheetName, int itemsPerSheet, CancellationToken cancellationToken = default)
            where T : class, new();

        /// <summary>
        /// 将数据集合写入到表格的指定列并返回流。
        /// </summary>
        /// <param name="data"></param>
        /// <param name="firstRowNumber">起始行号，从1开始</param>
        /// <param name="columnNumber">列号，起始值为1</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">数据对象类型</typeparam>
        /// <returns></returns>
        Task<Stream> WriteAsync<T>(IEnumerable<T> data, int firstRowNumber, int columnNumber, string sheetName, Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// 将数据集合写入到表格的指定列并返回流
        /// </summary>
        /// <param name="dataFactory"></param>
        /// <param name="firstRowNumber">起始行号，从1开始</param>
        /// <param name="columnNumber">列号，起始值为1</param>
        /// <param name="sheetName">表格名称，不指定取第一个</param>
        /// <param name="valueConvert">值转换方法</param>
        /// <param name="cancellationToken"></param>
        /// <typeparam name="T">数据对象类型</typeparam>
        /// <returns></returns>
        Task<Stream> WriteAsync<T>(Func<Task<IEnumerable<T>>> dataFactory, int firstRowNumber, int columnNumber, string sheetName, Func<T, CultureInfo, object> valueConvert = null, CancellationToken cancellationToken = default);

        #endregion
    }
}