using System;
using System.Collections.Generic;
using System.Linq;

namespace Nerosoft.Powersheet
{
    public class SheetWriteOptions : SheetHandleOptions
    {
        private readonly List<string> _ignoreNames = new();

        /// <summary>
        /// 获取忽略的属性或DataTable列
        /// </summary>
        public IEnumerable<string> IgnoreNames => _ignoreNames;

        public SheetColumnMapProfile GetMapProfile(string name)
        {
            return Mapping.FirstOrDefault(t => t.Name == name);
        }

        /// <summary>
        /// 添加忽略的属性
        /// </summary>
        /// <param name="names"></param>
        public void IgnoreName(params string[] names)
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