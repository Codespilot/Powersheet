using NPOI.SS.UserModel;
using System;

namespace Nerosoft.Powersheet.Npoi
{
    /// <summary>
    /// 基于NPOI实现表格解析
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

        private static int GetHeaderRowNumber(ISheet sheet, int defaultHeaderRowNumber = 1)
        {
            if (sheet.FirstRowNum == 0)
            {
                return defaultHeaderRowNumber - 1;
            }

            if (sheet.FirstRowNum == 1)
            {
                return defaultHeaderRowNumber;
            }

            if (defaultHeaderRowNumber > sheet.FirstRowNum)
            {
                return defaultHeaderRowNumber;
            }

            return sheet.FirstRowNum;
        }
    }
}