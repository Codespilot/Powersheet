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
        private readonly List<string> _ignoreNames = new();

        /// <summary>
        /// 获取或设置忽略的表格列
        /// </summary>
        public override IEnumerable<string> IgnoreNames => _ignoreNames;

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
        /// <param name="names"></param>
        public override void IgnoreName(params string[] names)
        {
            if (names == null || names.Length < 1)
            {
                return;
            }

            foreach (var name in names)
            {
                if (_ignoreNames.Contains(name, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                _ignoreNames.Add(name);
            }
        }
    }
}