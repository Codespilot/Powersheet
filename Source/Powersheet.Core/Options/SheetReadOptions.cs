using System;
using System.Collections.Generic;
using System.Linq;

namespace Nerosoft.Powersheet
{
    public class SheetReadOptions : SheetHandleOptions
    {
        private readonly List<string> _ignoreColumns = new();

        /// <summary>
        /// 获取或设置忽略的表格列
        /// </summary>
        public IEnumerable<string> IgnoreColumns => _ignoreColumns;

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