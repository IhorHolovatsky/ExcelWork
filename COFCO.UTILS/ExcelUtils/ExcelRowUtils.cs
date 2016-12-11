﻿using System;
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

        public static CofcoRowModel CopyRow(IRow inputRow, ExcelInputInfo inputInfo)
        {
            int idValue;
            var idColumnValue = inputRow.GetCell(ExcelConstants.HIDDEN_ID_COLUMN_INDEX).GetCellValue();
            
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

            if (int.TryParse(idColumnValue, out idValue))
            {
                cofcoModel.Id = idValue;
            }

            return cofcoModel;
        }

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

            outputRow.CreateCell(ExcelConstants.HIDDEN_ID_COLUMN_INDEX, CellType.Numeric)
                     .SetCellValue(rowModel.Id);
        }
    }
}
