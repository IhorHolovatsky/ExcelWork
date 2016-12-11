using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using COFCO.SharedEntities.Constants;
using COFCO.SharedEntities.Models;
using NPOI.SS.UserModel;

namespace COFCO.UTILS.ExcelUtils
{
    public class ExcelRowUtils
    {
        /// <summary>
        /// Copy cell data by from inputRow to outputRow
        /// </summary>
        /// <param name="columnIndexes">Indexes of columns of input row</param>
        public static void CopyRow(IRow inputRow, IRow outputRow, IEnumerable<int> columnIndexes)
        {
            var i = 0;
            foreach (var index in columnIndexes)
            {
                var inputCell = inputRow.GetCell(index);
                var outputCell = outputRow.CreateCell(i, CellType.String);

                outputCell.SetCellValue(inputCell.GetCellValue());
                i++;
            }
        }

        /// <summary>
        /// Copy Excel row data to CofcoRowModel
        /// </summary>
        /// <returns></returns>
        public static CofcoRowModel CopyRow(IRow inputRow, ExcelInputInfo inputInfo)
        {
            // -1 because, we start from 0
            var idColumnValue = inputRow.GetCell(inputRow.LastCellNum - 1).GetCellValue();
            
            var cofcoModel = new CofcoRowModel
            {
                Port = inputRow.GetCell(inputInfo.Port).GetCellValue(),
                Supplier = inputRow.GetCell(inputInfo.Supplier).GetCellValue(),
                Product = inputRow.GetCell(inputInfo.Product).GetCellValue(),
                Quantity = inputRow.GetCell(inputInfo.Quantity).GetCellValue(),
                Date = inputRow.GetCell(inputInfo.Date).GetCellValue(),
                VehicleNumber = inputRow.GetCell(inputInfo.VehicleNumber).GetCellValue(),
                TTNNumber = inputRow.GetCell(inputInfo.TTNNumber).GetCellValue(),
                Contract = inputRow.GetCell(inputInfo.Contract).GetCellValue()
            };
            cofcoModel.Id = idColumnValue;
            
            return cofcoModel;
        }

        /// <summary>
        /// Write Excel row with data from rowModel, with additional HiddenCell with Id
        /// </summary>
        public static void WriteRowWithHiddenId(ISheet sheet, IRow outputRow, CofcoRowModel rowModel)
        {
            WriteRow(outputRow, rowModel);

            outputRow.CreateCell(9, CellType.Numeric)
                     .SetCellValue(rowModel.Id);

            sheet.SetColumnHidden(9, true);
        }

        /// <summary>
        /// Write Excel row with data from rowModel
        /// </summary>
        /// <param name="outputRow">destination row</param>
        /// <param name="rowModel">row data</param>
        public static void WriteRow(IRow outputRow, CofcoRowModel rowModel)
        {
            outputRow.CreateCell(0, CellType.String)
                     .SetCellValue(rowModel.Port);
            outputRow.CreateCell(1, CellType.String)
                     .SetCellValue(rowModel.Supplier);
            outputRow.CreateCell(2, CellType.String)
                     .SetCellValue(rowModel.Product);

            int quantity;
            if (int.TryParse(rowModel.Quantity, out quantity))
            {
                outputRow.CreateCell(3, CellType.Numeric)
                         .SetCellValue(quantity);
            }
            else 
            {
                outputRow.CreateCell(3, CellType.String)
                         .SetCellValue(rowModel.Quantity);
            }
          
            outputRow.CreateCell(4, CellType.String)
                     .SetCellValue(rowModel.Date);
            outputRow.CreateCell(5, CellType.String)
                     .SetCellValue(rowModel.VehicleNumber);
            outputRow.CreateCell(6, CellType.String)
                     .SetCellValue(rowModel.TTNNumber);
            outputRow.CreateCell(7, CellType.String)
                     .SetCellValue(rowModel.Contract);
        }
    }
}
