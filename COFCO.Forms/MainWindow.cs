using System;
using System.Collections.Generic;
using System.Windows.Forms;
using COFCO.Forms.Helpers;
using COFCO.SharedEntities.Models;
using COFCO.UTILS.Extensions;

namespace COFCO.Forms
{
    public partial class MainWindow : Form
    {
        public static ExcelInputInfo ExcelInputInfoModel = new ExcelInputInfo();

        public static List<int> SupplierContractsOutputList = new List<int>();

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
                ExcelInputInfoModel.Port = tbPort.Text.ParseToInt().Value;
                ExcelInputInfoModel.Supplier = tbSupplier.Text.ParseToInt().Value;
                ExcelInputInfoModel.Product = tbProduct.Text.ParseToInt().Value;
                ExcelInputInfoModel.Quantity = tbQuantity.Text.ParseToInt().Value;
                ExcelInputInfoModel.Date = tbDate.Text.ParseToInt().Value;
                ExcelInputInfoModel.VehicleNumber = tbVehicleNumber.Text.ParseToInt().Value;
                ExcelInputInfoModel.TTNNumber = tbTTNNumber.Text.ParseToInt().Value;
                ExcelInputInfoModel.Contract = tbContact.Text.ParseToInt().Value;
                ExcelInputInfoModel.SheetNumber = tbSheetNumber.Text.ParseToInt().Value;
                ExcelInputInfoModel.StartRowNumber = tbStartRowNumber.Text.ParseToInt().Value;
            }
            catch (Exception ex)
            {
                ShowMessageBoxWithError();
            }


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

    }
}
