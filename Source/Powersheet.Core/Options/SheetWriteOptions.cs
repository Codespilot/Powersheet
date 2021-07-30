using System;
using System.Collections.Generic;
using System.Linq;

namespace Nerosoft.Powersheet
{
    /// <summary>
    /// 表格写入配置选项
    /// </summary>
    public class SheetWriteOptions : SheetHandleOptions
    {
        private readonly List<string> _ignoreNames = new();

        /// <summary>
        /// 获取忽略的属性或DataTable列
        /// </summary>
        public override IEnumerable<string> IgnoreNames => _ignoreNames;

        /// <summary>
        /// 标题单元格样式
        /// </summary>
        public CellStyle HeaderStyle { get; set; }

        /// <summary>
        /// 数据单元格样式
        /// </summary>
        public CellStyle BodyStyle { get; set; }

        /// <summary>
        /// 根据对象属性名/DataTable列名获取映射配置
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public SheetColumnMapProfile GetMapProfile(string name)
        {
            return Mapping.FirstOrDefault(t => t.Name == name);
        }

        /// <summary>
        /// 添加忽略的属性
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