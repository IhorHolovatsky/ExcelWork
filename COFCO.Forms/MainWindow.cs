using System.Windows.Forms;
using COFCO.Forms.Helpers;
using COFCO.SharedEntities.Models;
using COFCO.UTILS.Extensions;

namespace COFCO.Forms
{
    public partial class MainWindow : Form
    {
        public static ExcelInputInfo ExcelInputInfoModel = new ExcelInputInfo();

        public MainWindow()
        {
            InitializeComponent();
        }

        #region Buttons
        private void btnCreateTempExcel_Click(object sender, System.EventArgs e)
        {
            //ToDo: validation

            ExcelInputInfoModel.Port = tbPort.Text.ParseToInt();
            ExcelInputInfoModel.Supplier = tbSupplier.Text.ParseToInt();
            ExcelInputInfoModel.Product = tbProduct.Text.ParseToInt();
            ExcelInputInfoModel.Quantity = tbQuantity.Text.ParseToInt();
            ExcelInputInfoModel.Date = tbProduct.Text.ParseToInt();
            ExcelInputInfoModel.VehicleNumber = tbProduct.Text.ParseToInt();
            ExcelInputInfoModel.TTNNumber = tbProduct.Text.ParseToInt();
            ExcelInputInfoModel.Contract = tbProduct.Text.ParseToInt();
            ExcelInputInfoModel.SheetNumber = tbProduct.Text.ParseToInt().Value;
            ExcelInputInfoModel.StartRowNumber = tbProduct.Text.ParseToInt().Value;
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



    }
}
