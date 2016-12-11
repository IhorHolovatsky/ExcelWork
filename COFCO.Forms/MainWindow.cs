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
        private readonly ExcelService  _excelService = new ExcelService();
        private Excel.Worksheet _contractsWorksheet;
        private Excel.Application _excelapp;

        public static ExcelInputInfo ExcelInputInfoModel = new ExcelInputInfo();

        public static List<int> SupplierContractsOutputList;
        
        public MainWindow()
        {
            InitializeComponent();
        }

        #region Buttons
        private void btnCreateTempExcel_Click(object sender, System.EventArgs e)
        {
            try
            {
                ExcelInputInfoModel.Port = tbPort.Text.ParseToInt().Value - 1;
                ExcelInputInfoModel.Supplier = tbSupplier.Text.ParseToInt().Value - 1;
                ExcelInputInfoModel.Product = tbProduct.Text.ParseToInt().Value - 1;
                ExcelInputInfoModel.Quantity = tbQuantity.Text.ParseToInt().Value - 1;
                ExcelInputInfoModel.Date = tbDate.Text.ParseToInt().Value - 1;
                ExcelInputInfoModel.VehicleNumber = tbVehicleNumber.Text.ParseToInt().Value - 1;
                ExcelInputInfoModel.TTNNumber = tbTTNNumber.Text.ParseToInt().Value - 1;
                ExcelInputInfoModel.Contract = tbContact.Text.ParseToInt().Value - 1;
                ExcelInputInfoModel.SheetNumber = tbSheetNumber.Text.ParseToInt().Value - 1;
                ExcelInputInfoModel.StartRowNumber = tbStartRowNumber.Text.ParseToInt().Value - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Constants.ParamsInputErrorMessage, Constants.ErrorMessage);
                return;
            }

            if (String.IsNullOrEmpty(ExcelInputInfoModel.InputFilePath) ||
                String.IsNullOrEmpty(ExcelInputInfoModel.OutputTempFolderPath))
            {
                MessageBox.Show(this, Constants.InputFileAndDirectoryExistanceErrorMessage, Constants.ErrorMessage);
                return;
            }

            try
            {
                SupplierContractsOutputList = _excelService.CreateTempExcelFile(ExcelInputInfoModel);
            }
            catch (Exception ex)
            {
                FormsLogger.FormsLoggerInstance.Error(ex);
                MessageBox.Show(this, Constants.InputExcelErrorMessage, Constants.ErrorMessage);
                return;
            }

            _excelapp = new Excel.Application
            {
                Visible = true
            };

            var excelappworkbook = _excelapp.Workbooks.Open(Path.Combine(ExcelInputInfoModel.OutputTempFolderPath,FileContants.TEMP_EXCEL_FILE_NAME),
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing);

            var excelsheets = excelappworkbook.Worksheets;
            _contractsWorksheet = (Excel.Worksheet)excelsheets.Item[1];

            _contractsWorksheet.Change += ExcelworksheetOnChange;
            
        }
        
        private void btnCreateTemplates_Click(object sender, System.EventArgs e)
        {
            if (String.IsNullOrEmpty(ExcelInputInfoModel.InputFilePath) )
            {
                MessageBox.Show(this, Constants.InputFileExistanceErrorMessage, Constants.ErrorMessage);
                return;
            }

            if (String.IsNullOrEmpty(ExcelInputInfoModel.TempExcelFilePath) ||
                String.IsNullOrEmpty(ExcelInputInfoModel.OutputTemplateFolderPath))
            {
                MessageBox.Show(this, Constants.TempFileAndDirectoryExistanceErrorMessage, Constants.ErrorMessage);
                return;
            }

            _excelService.FillExcelWithMissedColumns(ExcelInputInfoModel);
            
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
        }

        private void btnChooseOutputTempFolder_Click(object sender, System.EventArgs e)
        {
            var filePath = DialogHelper.OpenChooseFolderDialog(tbOutputFolderPath);

            if (!string.IsNullOrEmpty(filePath))
            {
                ExcelInputInfoModel.OutputTempFolderPath = filePath;
            }
        }

        private void btnChooseExcelFile_Click(object sender, System.EventArgs e)
        {
            var filePath = DialogHelper.OpenChooseFileDialog(tbExcelFilePath);

            if (!string.IsNullOrEmpty(filePath))
            {
                ExcelInputInfoModel.TempExcelFilePath = filePath;
            }

        }

        private void btnChooseTemplateFolderPath_Click(object sender, System.EventArgs e)
        {
            var filePath = DialogHelper.OpenChooseFolderDialog(tbOutputTemplateFolder);

            if (!string.IsNullOrEmpty(filePath))
            {
                ExcelInputInfoModel.OutputTemplateFolderPath = filePath;
            }
        }
        #endregion

        #region Excel Event Handlers
        private void ExcelworksheetOnChange(Excel.Range target)
        {
            var targetRowAdress = Convert.ToInt32(ContractBL.GetRowAdressByRange(target));

            if (SupplierContractsOutputList.Contains(targetRowAdress))
            {
                return;
            }

            ContractBL.FeelContractsSummary(SupplierContractsOutputList, _contractsWorksheet);

        }
        #endregion




    }
}
