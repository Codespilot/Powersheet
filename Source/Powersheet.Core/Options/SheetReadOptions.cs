using System;
using System.Collections.Generic;
using System.Linq;

namespace Nerosoft.Powersheet
{
    /// <summary>
    /// 表格读取配置选项
    /// </summary>
    public class SheetReadOptions : SheetHandleOptions
    {
        private readonly List<string> _ignoreColumns = new();

        /// <summary>
        /// 获取或设置忽略的表格列
        /// </summary>
        public IEnumerable<string> IgnoreColumns => _ignoreColumns;

        /// <summary>
        /// 根据表格列名获取映射配置
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public SheetColumnMapProfile GetMapProfile(string name)
        {
            return Mapping.FirstOrDefault(t => t.ColumnName == name);
        }

        /// <summary>
        /// 添加忽略的表格列
        /// </summary>
        /// <param name="columnsName"></param>
        public void IgnoreColumn(params string[] columnsName)
        {
            if (columnsName == null || columnsName.Length < 1)
            {
                return;
            }

            foreach (var name in columnsName)
            {
                if (_ignoreColumns.Contains(name, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                _ignoreColumns.Add(name);
            }
        }
    }
}