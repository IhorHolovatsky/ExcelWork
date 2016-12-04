using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

namespace COFCO.UTILS.ExcelUtils
{
    public static class ExcelCellUtils
    {
        public static string GetCellValue(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                default:
                    return string.Empty;
            }
        }
    }
}
