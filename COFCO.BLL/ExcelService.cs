﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using COFCO.SharedEntities.Constants;
using COFCO.SharedEntities.Models;
using COFCO.UTILS.ExcelUtils;
using COFCO.UTILS.Extensions;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace COFCO.BLL
{
    public class ExcelService
    {
        private static XSSFWorkbook inputExcel { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelInputInfo"></param>
        public List<int> CreateTempExcelFile(ExcelInputInfo excelInputInfo)
        {
            if (string.IsNullOrWhiteSpace(excelInputInfo.InputFilePath)
                || string.IsNullOrWhiteSpace(excelInputInfo.OutputTempFolderPath))
            {
                throw new Exception("Input file path or Output temp folder path is empty!");
            }
            var totalSumRowIndexes = new List<int>();

            var inputFileExtension = Path.GetExtension(excelInputInfo.InputFilePath);
            var isXlsx = string.Equals(inputFileExtension, FileContants.XLSX);

            if (isXlsx)
            {
                #region Logic for XLSX


                XSSFWorkbook outputTempExcel;


                try
                {
                    using (var file = new FileStream(excelInputInfo.InputFilePath, FileMode.Open, FileAccess.Read))
                    {
                        inputExcel = new XSSFWorkbook(file);
                    }

                    outputTempExcel = CreateEmptyTempExcel();
                }
                catch (OfficeXmlFileException e)
                {
                    throw new Exception("Invalid excel extension");
                }
                catch (IOException e)
                {
                    throw e;
                }

                FillExcelWithHeaders(outputTempExcel);

                var inputSheet = inputExcel.GetSheetAt(excelInputInfo.SheetNumber);
                var outputSheet = outputTempExcel.GetSheetAt(0);

                var dataList = new List<CofcoRowModel>(inputSheet.LastRowNum - excelInputInfo.StartRowNumber);


                for (var rowIndex = excelInputInfo.StartRowNumber; rowIndex <= inputSheet.LastRowNum; rowIndex++)
                {
                    var inputRow = inputSheet.GetRow(rowIndex);

                    //set hidden Column for ID
                    inputRow.CreateCell(ExcelConstants.HIDDEN_ID_COLUMN_INDEX, CellType.Numeric)
                            .SetCellValue(rowIndex);
                    inputSheet.SetColumnHidden(ExcelConstants.HIDDEN_ID_COLUMN_INDEX, true);

                    dataList.Add(ExcelRowUtils.CopyRow(inputRow, excelInputInfo));
                }

                dataList = dataList.OrderBy(d => d.Supplier)
                                   .ToList();



                var previousSupplier = string.Empty;
                //1 - because first  row -> with headers
                var i = 1;
                var supplierSum = 0;
                foreach (var cofcoData in dataList)
                {
                    if (cofcoData.Supplier != string.Empty
                        && previousSupplier != string.Empty
                        && cofcoData.Supplier != previousSupplier)
                    {
                        CreateSummaryRows(outputSheet, ref i, ref supplierSum, ref totalSumRowIndexes);
                    }

                    var newRow = outputSheet.CreateRow(i);
                    ExcelRowUtils.WriteRow(newRow, cofcoData);

                    var quantity = cofcoData.Quantity.ParseToInt();
                    if (quantity.HasValue)
                    {
                        supplierSum += quantity.Value;
                    }

                    previousSupplier = cofcoData.Supplier;
                    i++;
                }

                //For last supplier
                CreateSummaryRows(outputSheet, ref i, ref supplierSum, ref totalSumRowIndexes);

                SaveExcel(outputTempExcel, excelInputInfo);

                inputExcel.Close();
                outputTempExcel.Close();

                return totalSumRowIndexes;

                #endregion
            }
            else
            {

            }

            return totalSumRowIndexes;
        }

        public void SaveTemplatesByDate(ExcelInputInfo excelInputInfo)
        {

        }

        public void SaveTemplatesBySupplier(ExcelInputInfo excelInputInfo)
        {

        }

        #region private methods
        private XSSFWorkbook CreateEmptyTempExcel()
        {
            var hssfworkbook = new XSSFWorkbook();

            //here, we must insert at least one sheet to the workbook. otherwise, Excel will say 'data lost in file'
            hssfworkbook.CreateSheet("Постачальники");

            return hssfworkbook;
        }

        private void FillExcelWithHeaders(XSSFWorkbook workbook)
        {
            var outputSheet = workbook.GetSheetAt(0);
            var outputHeaderRow = outputSheet.CreateRow(0);

            var cofcoHeaderModel = new CofcoRowModel
            {
                Port = "Порт",
                Supplier = "Постачальник",
                Product = "Продукт",
                Quantity = "Кількість",
                Date = "Дата",
                VehicleNumber = "Номер машини",
                TTNNumber = "Номер ТТН",
                Contract = "Контракт(Старий)"
            };

            ExcelRowUtils.WriteRow(outputHeaderRow, cofcoHeaderModel);

            outputHeaderRow.CreateCell(8, CellType.String)
                           .SetCellValue("Контракт");

            foreach (var i in Enumerable.Range(0, 9))
            {
                outputSheet.SetColumnWidth(i, 6000);
            }

        }

        private string SaveExcel(XSSFWorkbook workbook, ExcelInputInfo excelInputInfo)
        {
            var filePath = Path.Combine(excelInputInfo.OutputTempFolderPath, FileContants.TEMP_EXCEL_FILE_NAME);
            //Write the stream data of workbook to the root directory
            var file = new FileStream(filePath, FileMode.Create);
            workbook.Write(file);
            file.Close();

            return filePath; ;
        }

        private void CreateSummaryRows(ISheet outputSheet, ref int rowNumber, ref int supplierSum, ref List<int> totalSumRowIndexes)
        {
            var supplierSumRow = outputSheet.CreateRow(rowNumber);
            supplierSumRow.CreateCell(0)
                          .SetCellValue("Cума по постачальнику:");
            supplierSumRow.CreateCell(1, CellType.Numeric)
                          .SetCellValue(supplierSum.ToString());

            supplierSum = 0;
            rowNumber++;

            var summaryRow = outputSheet.CreateRow(rowNumber);

            summaryRow.CreateCell(0)
                      .SetCellValue("Сума по контракту:");
            rowNumber++;

            totalSumRowIndexes.Add(rowNumber);
        }

        private List<IRow> GetPrimaryInputRows(XSSFWorkbook inputExcel, ExcelInputInfo excelInputInfo)
        {
            var returnValue = new List<IRow>();

            var sheet = inputExcel.GetSheetAt(excelInputInfo.SheetNumber);

            for (var i = excelInputInfo.StartRowNumber; i < sheet.LastRowNum; i++)
            {
                returnValue.Add(sheet.GetRow(i));
            }


            return returnValue;
        }


        #endregion
    }
}
