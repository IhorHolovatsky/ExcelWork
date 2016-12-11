using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using COFCO.SharedEntities.Constants;
using COFCO.SharedEntities.Models;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace COFCO.UTILS.ExcelUtils
{
    public class ExcelSheetUtils
    {
        public static List<Dictionary<int, string>> CopyAllRows(ISheet inputSheet, ExcelInputInfo excelInputInfo)
        {
            var inputValues = new List<Dictionary<int, string>>(inputSheet.LastRowNum - excelInputInfo.StartRowNumber);

            for (var i = excelInputInfo.StartRowNumber; i <= inputSheet.LastRowNum; i++)
            {
                var inputRow = inputSheet.GetRow(i);

                if (inputRow == null)
                    continue;

                var rowValues = new Dictionary<int, string>();

                for (var j = inputRow.FirstCellNum; j <= inputRow.LastCellNum; j++)
                {
                    if (IsNeededColumn(excelInputInfo, j))
                    {
                        rowValues.Add(j, inputRow.GetCell(j).GetCellValue());
                    }
                }

                //rowValues = rowValues.Where(node => !string.IsNullOrEmpty(node.Value))
                //                     .ToDictionary(node => node.Key, node => node.Value);

                //Remove cell with hidden id's
                inputRow.RemoveCell(inputRow.GetCell(inputRow.LastCellNum - 1));
                inputValues.Add(rowValues);
            }

            return inputValues;
        }

        #region Private Methods

        private static bool IsNeededColumn(ExcelInputInfo excelInputInfo, int columnIndex)
        {
            return columnIndex != excelInputInfo.Contract
                   && columnIndex != excelInputInfo.Date
                   && columnIndex != excelInputInfo.Port
                   && columnIndex != excelInputInfo.Quantity
                   && columnIndex != excelInputInfo.Product
                   && columnIndex != excelInputInfo.TTNNumber
                   && columnIndex != excelInputInfo.VehicleNumber
                   && columnIndex != excelInputInfo.Supplier;
        }
    }
    #endregion
}
