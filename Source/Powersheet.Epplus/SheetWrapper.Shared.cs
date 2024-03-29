using System;

namespace Nerosoft.Powersheet.Epplus
{
    /// <summary>
    /// 基于EPPLUS实现表格解析
    /// </summary>
    public partial class SheetWrapper : SheetWrapperBase
    {
        private static readonly Lazy<SheetWrapper> _instance = new();

        /// <summary>
        /// 获取默认<see cref="ISheetWrapper"/>实例
        /// </summary>
        public static ISheetWrapper Default => _instance.Value;
        
        /// <summary>
        /// 获取新的<see cref="ISheetWrapper"/>实例
        /// </summary>
        public static ISheetWrapper New => new SheetWrapper();
    }
}