using System;
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
        private List<int> _totalSumRowIndexes;

        /// <summary>
        ///  Read xlsx file
        /// </summary>
        public XSSFWorkbook ReadExcelFile(string filePath)
        {
            XSSFWorkbook returnValue;

            try
            {
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    returnValue = new XSSFWorkbook(file);
                }
            }
            catch (OfficeXmlFileException e)
            {
                throw new Exception("Invalid excel extension");
            }
            catch (IOException e)
            {
                throw e;
            }

            return returnValue;
        }

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

                inputExcel = ReadExcelFile(excelInputInfo.InputFilePath);

                var outputTempExcel = CreateEmptyTempExcel();

                FillExcelWithCustomHeaders(outputTempExcel);

                var inputSheet = inputExcel.GetSheetAt(excelInputInfo.SheetNumber);
                var outputSheet = outputTempExcel.GetSheetAt(0);

                var dataList = new List<CofcoRowModel>(inputSheet.LastRowNum - excelInputInfo.StartRowNumber);


                for (var rowIndex = excelInputInfo.StartRowNumber; rowIndex <= inputSheet.LastRowNum; rowIndex++)
                {
                    var inputRow = inputSheet.GetRow(rowIndex);
                    
                    if (inputRow == null)
                        continue;

                    //set hidden Column for ID
                    inputRow.CreateCell(inputRow.LastCellNum + 1, CellType.Numeric)
                            .SetCellValue($"{Guid.NewGuid()}|{rowIndex}");
                    inputSheet.SetColumnHidden(inputRow.LastCellNum, true);

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

                    ExcelRowUtils.WriteRowWithHiddenId(outputSheet, newRow, cofcoData);


                    var quantity = cofcoData.Quantity.ParseToInt();
                    if (quantity.HasValue)
                    {
                        supplierSum += quantity.Value;
                    }

                    previousSupplier = cofcoData.Supplier;
                    i++;
                }

                _totalSumRowIndexes = totalSumRowIndexes;

                //For last supplier
                CreateSummaryRows(outputSheet, ref i, ref supplierSum, ref totalSumRowIndexes);

                SaveExcel(outputTempExcel, excelInputInfo.OutputTempFolderPath);

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

        public XSSFWorkbook FillExcelWithMissedColumns(ExcelInputInfo excelInputInfo)
        {
            if (inputExcel == null)
            {
                if (string.IsNullOrEmpty(excelInputInfo.InputFilePath))
                {
                    throw new ArgumentNullException("Missed Input File");
                }

                inputExcel = ReadExcelFile(excelInputInfo.InputFilePath);
            }
            var excelWithFilledContacts = ReadExcelFile(excelInputInfo.TempExcelFilePath);

            var inputSheet = inputExcel.GetSheetAt(excelInputInfo.SheetNumber);
            var outputSheet = excelWithFilledContacts.GetSheetAt(0);

            var inputValues = ExcelSheetUtils.CopyAllRows(inputSheet, excelInputInfo);

            //I Guess that first row is row with headers
            FillExcelWithHeaders(inputSheet.GetRow(0), outputSheet, inputValues.First()
                                                                         .Select(node => node.Key)
                                                                         .ToList());

            //i = 1, because first row -> with headers
            for (var i = 1; i <= outputSheet.LastRowNum; i++)
            {
                var outputRow = outputSheet.GetRow(i);
                var hiddenCell = outputRow.GetCell(9);
                var hiddenCellValue = hiddenCell.GetCellValue();


                //Skip summary rows
                if (string.IsNullOrEmpty(hiddenCellValue))
                {
                    continue;
                }

                outputRow.RemoveCell(hiddenCell);
                outputSheet.SetColumnHidden(9, false);

                var missedValues = inputValues.First(node => node.Values.Contains(hiddenCellValue))
                                              .Select(node => node.Value)
                                              .ToList();

                missedValues.Remove(hiddenCellValue);

                var lastCellNumber = 9;

                foreach (string value in missedValues)
                {
                    outputRow.CreateCell(lastCellNumber, CellType.String)
                             .SetCellValue(value);
                    lastCellNumber ++;
                }
            }

            //SaveExcel(excelWithFilledContacts, excelInputInfo.OutputTemplateFolderPath);
            SaveTemplatesByDate(excelInputInfo ,excelWithFilledContacts);
            SaveTemplatesBySupplier(excelInputInfo, excelWithFilledContacts);

            inputExcel.Close();
            excelWithFilledContacts.Close();

            return excelWithFilledContacts;
        }

        public void SaveTemplatesByDate(ExcelInputInfo excelInputInfo, XSSFWorkbook workbook)
        {
            var currentDate = DateTime.Today;

            var currentDateFolderName = currentDate.Day + "." + currentDate.Month + "." + currentDate.Year;
            
            var pathWithDate = Path.Combine(excelInputInfo.OutputTemplateFolderPath, "ByDate",
               currentDateFolderName);

            SaveExcelTemplateByDate(workbook, pathWithDate);

        }

        public void SaveTemplatesBySupplier(ExcelInputInfo excelInputInfo, XSSFWorkbook workbook)
        {
            var inputWorksheet = workbook.GetSheetAt(0);
            var headersRow = inputWorksheet.GetRow(0);
            
            var lastIterationNumber = 0;

            var inputRowIterationNumber = 1;

            foreach (var totalSumRowIndex in _totalSumRowIndexes)
            {
                
                var templateWorkbook = new XSSFWorkbook();

                templateWorkbook.CreateSheet("Звіт");

                var outputSheet = templateWorkbook.GetSheetAt(0);
                var outputHeaderRow = outputSheet.CreateRow(0);
                

                for (int i = 2; i < headersRow.LastCellNum; i++)
                {
                    outputHeaderRow.CreateCell(i-2, CellType.String)
                     .SetCellValue(headersRow.GetCell(i).StringCellValue);
                }



                var rowIndex = 1;
                for (int j = lastIterationNumber + 1; j <= totalSumRowIndex-1; j++)
                {
                    var currentRow = outputSheet.CreateRow(rowIndex);
                    var inputRow = inputWorksheet.GetRow(inputRowIterationNumber);
                    rowIndex++;
                    inputRowIterationNumber++;
                    
                    if (totalSumRowIndex - j <= 1)
                    {
                        for (int i = 0; i < headersRow.LastCellNum; i++)
                        {
                            currentRow.CreateCell(i, CellType.String)
                                .SetCellValue(inputRow.GetCell(i).GetCellValue());
                        }
                    }
                    else
                    {
                        for (int i = 2; i < headersRow.LastCellNum; i++)
                        {
                            currentRow.CreateCell(i - 2, CellType.String)
                                .SetCellValue(inputRow.GetCell(i).GetCellValue());
                        }
                    }

                    
                }

                SaveExcelTemplateBySupplier(templateWorkbook, excelInputInfo.OutputTemplateFolderPath);

                lastIterationNumber = totalSumRowIndex;
            }
        }

        #region private methods
        private XSSFWorkbook CreateEmptyTempExcel()
        {
            var hssfworkbook = new XSSFWorkbook();

            //here, we must insert at least one sheet to the workbook. otherwise, Excel will say 'data lost in file'
            hssfworkbook.CreateSheet("Постачальники");

            return hssfworkbook;
        }

        private void FillExcelWithCustomHeaders(XSSFWorkbook workbook)
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

        private void FillExcelWithHeaders(IRow inputHeaderRow, ISheet outputSheet, List<int> columnIndexes)
        {
            var outputHeaderRow = outputSheet.GetRow(0);
            
            foreach (var columnIndex in columnIndexes)
            {
                outputHeaderRow.CreateCell(outputHeaderRow.LastCellNum)
                               .SetCellValue(inputHeaderRow.GetCell(columnIndex)
                                                           .GetCellValue());
            }
        }

        private string SaveExcel(XSSFWorkbook workbook, string filePath)
        {
            var filePathWithName = Path.Combine(filePath, FileContants.TEMP_EXCEL_FILE_NAME);
            //Write the stream data of workbook to the root directory
            var file = new FileStream(filePathWithName, FileMode.Create);
            workbook.Write(file);
            file.Close();

            return filePath; 
        }

        private string SaveExcelTemplateByDate(XSSFWorkbook workbook,string filePath)
        {
            Directory.CreateDirectory(filePath);

            var filePathWithName = Path.Combine(filePath, FileContants.TEMPLATE_EXCEL_FILE_NAME);
            //Write the stream data of workbook to the root directory
            var file = new FileStream(filePathWithName, FileMode.Create);
            workbook.Write(file);
            file.Close();

            return filePath; 
        }

        private string SaveExcelTemplateBySupplier(XSSFWorkbook workbook, string filePath)
        {
            var filePathWithName = Path.Combine(filePath, "ContractTempTEST.xlsx");
            //Write the stream data of workbook to the root directory
            var file = new FileStream(filePathWithName, FileMode.Create);
            workbook.Write(file);
            file.Close();

            return filePath;
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

        #endregion
    }
}
