using NPOI.SS.UserModel;

namespace Nerosoft.Powersheet.Npoi
{
    /// <summary>
    /// 基于NPOI实现表格解析
    /// </summary>
    public partial class SheetWrapper : SheetWrapperBase
    {
        private static int GetHeaderRowNumber(ISheet sheet, int defaultHeaderRowNumber = 1)
        {
            return sheet.FirstRowNum switch
            {
                0 => defaultHeaderRowNumber - 1,
                1 => defaultHeaderRowNumber,
                _ => defaultHeaderRowNumber
            };
        }
    }
}