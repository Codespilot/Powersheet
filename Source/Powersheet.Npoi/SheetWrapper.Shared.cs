using NPOI.SS.UserModel;
using System;

namespace Nerosoft.Powersheet.Npoi
{
    /// <summary>
    /// 基于NPOI实现表格解析
    /// </summary>
    public partial class SheetWrapper : SheetWrapperBase
    {
        private static readonly Lazy<SheetWrapper> Instance = new();

        /// <summary>
        /// 获取默认<see cref="ISheetWrapper"/>实例
        /// </summary>
        public static ISheetWrapper Default => Instance.Value;

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