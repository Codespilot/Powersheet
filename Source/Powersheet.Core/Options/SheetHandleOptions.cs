using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;

namespace Nerosoft.Powersheet
{
    /// <summary>
    /// To be added.
    /// </summary>
    public abstract class SheetHandleOptions
    {
        private readonly List<SheetColumnMapProfile> _mapping = new();

        /// <summary>
        /// 获取或设置列头所在行号
        /// </summary>
        public int HeaderRowNumber { get; set; } = 1;

        /// <summary>
        /// 获取或设置首列起始位置
        /// </summary>
        public int FirstColumnNumber { get; set; } = 1;

        /// <summary>
        /// 获取或设置读取/写入的行数
        /// </summary>
        public int? RowCount { get; set; }

        /// <summary>
        /// 获取忽略的项（表格列名或对象属性名）
        /// </summary>
        public abstract IEnumerable<string> IgnoreNames { get; }

        /// <summary>
        /// 添加忽略的项（表格列名或对象属性名）
        /// </summary>
        /// <param name="names"></param>
        public abstract void IgnoreName(params string[] names);

        /// <summary>
        /// 获取列映射配置
        /// </summary>
        public IEnumerable<SheetColumnMapProfile> Mapping => _mapping;

        /// <summary>
        /// 添加映射配置
        /// </summary>
        /// <param name="option"></param>
        /// <returns>当前实例</returns>
        public SheetHandleOptions AddMapProfile(SheetColumnMapProfile option)
        {
            if (_mapping.Any(t => t.Name.Equals(option.Name)))
            {
                throw new DuplicateNameException($"Duplicate property/column name '{option.Name}' has already exists.");
            }

            if (_mapping.Any(t => t.ColumnName.Equals(option.ColumnName)))
            {
                throw new DuplicateNameException($"Duplicate sheet column name '{option.ColumnName}'.");
            }

            _mapping.Add(option);
            return this;
        }

        /// <summary>
        /// 添加映射配置
        /// </summary>
        /// <param name="name"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public SheetHandleOptions AddMapProfile(string name, string columnName)
        {
            return AddMapProfile(new SheetColumnMapProfile(name, columnName));
        }

        /// <summary>
        /// 添加映射配置
        /// </summary>
        /// <param name="name"></param>
        /// <param name="columnName"></param>
        /// <param name="valueConvert"></param>
        /// <returns></returns>
        public SheetHandleOptions AddMapProfile(string name, string columnName, Func<object, CultureInfo, object> valueConvert)
        {
            return AddMapProfile(new SheetColumnMapProfile(name, columnName, valueConvert));
        }

        /// <summary>
        /// 添加映射配置
        /// </summary>
        /// <param name="name"></param>
        /// <param name="columnName"></param>
        /// <param name="valueConverter"></param>
        /// <returns></returns>
        public SheetHandleOptions AddMapProfile(string name, string columnName, ICellValueConverter valueConverter)
        {
            return AddMapProfile(new SheetColumnMapProfile(name, columnName, valueConverter));
        }

        /// <summary>
        /// 添加映射配置
        /// </summary>
        /// <param name="name"></param>
        /// <param name="columnName"></param>
        /// <param name="valueConverterType"></param>
        /// <returns></returns>
        public SheetHandleOptions AddMapProfile(string name, string columnName, Type valueConverterType)
        {
            return AddMapProfile(new SheetColumnMapProfile(name, columnName, valueConverterType));
        }
    }
}