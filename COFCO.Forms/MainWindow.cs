using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using COFCO.BLL;
using COFCO.Forms.Helpers;
using COFCO.SharedEntities.Constants;
using COFCO.SharedEntities.Models;
using COFCO.UTILS.Extensions;
using Excel = Microsoft.Office.Interop.Excel;

namespace COFCO.Forms
{
    public partial class MainWindow : Form
    {
        public static ExcelInputInfo ExcelInputInfoModel = new ExcelInputInfo();

        public static List<int> SupplierContractsOutputList;

        public MainWindow()
        {
            InitializeComponent();
        }

        #region Buttons
        private void btnCreateTempExcel_Click(object sender, System.EventArgs e)
        {
            //ToDo: validation
            try
            {
                ExcelInputInfoModel.Port = tbPort.Text.ParseToInt().Value + 1;
                ExcelInputInfoModel.Supplier = tbSupplier.Text.ParseToInt().Value + 1;
                ExcelInputInfoModel.Product = tbProduct.Text.ParseToInt().Value + 1;
                ExcelInputInfoModel.Quantity = tbQuantity.Text.ParseToInt().Value + 1;
                ExcelInputInfoModel.Date = tbDate.Text.ParseToInt().Value + 1;
                ExcelInputInfoModel.VehicleNumber = tbVehicleNumber.Text.ParseToInt().Value + 1;
                ExcelInputInfoModel.TTNNumber = tbTTNNumber.Text.ParseToInt().Value + 1;
                ExcelInputInfoModel.Contract = tbContact.Text.ParseToInt().Value + 1;
                ExcelInputInfoModel.SheetNumber = tbSheetNumber.Text.ParseToInt().Value + 1;
                ExcelInputInfoModel.StartRowNumber = tbStartRowNumber.Text.ParseToInt().Value + 1;
            }
            catch (Exception ex)
            {
                ShowMessageBoxWithError();
                return;
            }

            SupplierContractsOutputList = new ExcelService().CreateTempExcelFile(ExcelInputInfoModel);
          
            var excelapp = new Excel.Application
            {
                Visible = true
            };

            var excelappworkbook = excelapp.Workbooks.Open(Path.Combine(ExcelInputInfoModel.OutputTempFolderPath,FileContants.TEMP_EXCEL_FILE_NAME),
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing);

            var excelsheets = excelappworkbook.Worksheets;
            var excelworksheet = (Excel.Worksheet)excelsheets.Item[1];

            excelworksheet.Change += target =>
            {
                int targetRowAdress = Convert.ToInt32(GetRowAdressByRange(target));

                if (SupplierContractsOutputList.Contains(targetRowAdress))
                {
                    return;
                }
                
                int lastIterationRowNumber = 1;

                foreach (var supplierRowNumber in SupplierContractsOutputList)
                {
                    var contractsDictionary = new Dictionary<string,double>();

                    for (int i = lastIterationRowNumber + 1; i < supplierRowNumber - 1; i++)
                    {
                        var contractCell = excelworksheet.Cells[i, 9];
                        var contractRange = excelworksheet.Range[contractCell, contractCell];
                        var contractValue = contractRange.Value2?.ToString();

                        var quantityCell = excelworksheet.Cells[i, 4];
                        var quantityRange = excelworksheet.Range[quantityCell, quantityCell];
                        var quantityValue = Convert.ToDouble(quantityRange.Value2);

                        if (contractValue != null)
                        {
                            if (contractsDictionary.ContainsKey(contractValue))
                            {
                                contractsDictionary[contractValue] = contractsDictionary[contractValue] + quantityValue;
                            }
                            else
                            {
                                contractsDictionary.Add(contractValue, quantityValue);
                            }
                        }
                    }

                    var outputString = String.Empty;

                    foreach (var item in contractsDictionary)
                    {
                        outputString += item.Key + " контракт : " + item.Value + ". ";
                    }

                    var outputAdress = "B" + supplierRowNumber;
                    var outputRange = excelworksheet.Range[outputAdress, outputAdress];
                    outputRange.Value2 = outputString;

                    lastIterationRowNumber = supplierRowNumber;
                }

            };
        }

        private void btnCreateTemplates_Click(object sender, System.EventArgs e)
        {

        }
        #endregion

        #region File/Folder dialogs

        private void btnChooseInputFile_Click(object sender, System.EventArgs e)
        {
            var inputFilePath = DialogHelper.OpenChooseFileDialog(tbInputFilePath);

            if (!string.IsNullOrEmpty(inputFilePath))
            {
                ExcelInputInfoModel.InputFilePath = inputFilePath;
            }
            else
            {
                //ToDo
            }
        }

        private void btnChooseOutputTempFolder_Click(object sender, System.EventArgs e)
        {
            var filePath = DialogHelper.OpenChooseFolderDialog(tbOutputFolderPath);

            if (!string.IsNullOrEmpty(filePath))
            {
                ExcelInputInfoModel.OutputTempFolderPath = filePath;
            }
            else
            {
                //ToDo
            }
        }

        private void btnChooseExcelFile_Click(object sender, System.EventArgs e)
        {
            var filePath = DialogHelper.OpenChooseFileDialog(tbExcelFilePath);

            if (!string.IsNullOrEmpty(filePath))
            {
                ExcelInputInfoModel.TempExcelFilePath = filePath;
            }
            else
            {
                //ToDo
            }

        }

        private void btnChooseTemplateFolderPath_Click(object sender, System.EventArgs e)
        {
            var filePath = DialogHelper.OpenChooseFolderDialog(tbOutputTemplateFolder);

            if (!string.IsNullOrEmpty(filePath))
            {
                ExcelInputInfoModel.OutputTemplateFolderPath = filePath;
            }
            else
            {
                //ToDo
            }
        }
        #endregion

        private void ShowMessageBoxWithError()
        {
            MessageBox.Show(this, "Перевірте правильність вводу. Всі колонки повинні бути заповнені та не має бути продубльованих рядків.");
        }

        private string GetRowAdressByRange(Excel.Range range)
        {
            var rangeAdress = range.Address;

            var adress = String.Empty;

            foreach (var chr in rangeAdress)
            {
                if (Char.IsDigit(chr))
                {
                    adress += chr;
                }
            }

            return adress;
        }

    }
}
