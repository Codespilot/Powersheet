using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;

namespace Nerosoft.Powersheet
{
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
        /// 获取列映射配置
        /// </summary>
        public IEnumerable<SheetColumnMapProfile> Mapping => _mapping;

        public void AddMapProfile(SheetColumnMapProfile option)
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
        }

        public void AddMapProfile(string name, string columnName)
        {
            AddMapProfile(new SheetColumnMapProfile(name, columnName));
        }

        public void AddMapProfile(string name, string columnName, Func<object, CultureInfo, object> valueConvert)
        {
            AddMapProfile(new SheetColumnMapProfile(name, columnName, valueConvert));
        }

        public void AddMapProfile(string name, string columnName, ICellValueConverter valueConverter)
        {
            AddMapProfile(new SheetColumnMapProfile(name, columnName, valueConverter));
        }
    }
}