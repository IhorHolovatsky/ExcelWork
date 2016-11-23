using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestApp
{
    public partial class Form1 : Form
    {

        private Excel.Application _excelapp;
        private Excel.Workbooks _excelappworkbooks;


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            _excelapp = new Excel.Application
            {
                Visible = true
            };
            
            _excelappworkbooks = _excelapp.Workbooks;

            var excelappworkbook = _excelapp.Workbooks.Open(GetExcelPath(),
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing);

            var excelsheets = excelappworkbook.Worksheets;
            //Получаем ссылку на лист 1
            var excelworksheet = (Excel.Worksheet)excelsheets.Item[Convert.ToInt32(textBoxSheet.Text)];
            //Выбираем ячейку для вывода A1/



            int firstColumn = Convert.ToInt32(column1textBox.Text);
            int secondColumn = Convert.ToInt32(column2textBox.Text);
            int thirdColumn = Convert.ToInt32(column3textBox.Text);



            var elementsList = new List<MainInfoModel>();

            //i = 2, to skip headers
            for (var i = 2; i <= excelworksheet.UsedRange.Rows.Count; i++)
            {
                var firstColumnValue = excelworksheet.Cells[i, firstColumn]?.Value2.ToString();
                var secondColumnValue = excelworksheet.Cells[i, secondColumn]?.Value2.ToString();
                var thirdColumnValue = excelworksheet.Cells[i, thirdColumn]?.Value2.ToString();

                elementsList.Add(new MainInfoModel()
                {
                    FirstString = firstColumnValue,
                    SecondString = secondColumnValue,
                    ThirdString = thirdColumnValue
                });
            }
           

            var outputExcelappWorkbook = _excelapp.Workbooks.Add();
            //var excelappworkbook = application.Workbooks.Open(@"E:\notjob\ExcelTestApp\aa.xlsx",
            //                   Type.Missing, Type.Missing, Type.Missing,
            // "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
            //  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //  Type.Missing, Type.Missing);
            //Если бы мы открыли несколько книг, то получили ссылку так
            //excelappworkbook=excelappworkbooks[1];
            //Получаем массив ссылок на листы выбранной книги
            var outputExcelSheets = outputExcelappWorkbook.Worksheets;
            //Получаем ссылку на лист 1
            var outputExcelWorksheet = (Excel.Worksheet)outputExcelSheets.get_Item(1);
            //Выбираем ячейку для вывода A1
            //var excelCells = excelworksheet.get_Range("A1", "A1");
            ////Выводим число
            //excelCells.Value2 = 10.5;

            var firstTitle = outputExcelWorksheet.get_Range("A1", "A1");
            //Выводим число
            firstTitle.Value2 = "FirstTitle";

            var secondTitle = outputExcelWorksheet.get_Range("B1", "B1");
            //Выводим число
            secondTitle.Value2 = "SecondTitle";

            var thirdTitle = outputExcelWorksheet.get_Range("C1", "C1");
            //Выводим число
            thirdTitle.Value2 = "ThirdTitle";

            var outexcelappworkbooks = _excelapp.Workbooks;
            var outexcelappworkbook = outexcelappworkbooks[2];
            outexcelappworkbook.SaveAs(@"E:\notjob\ExcelWork\aa.xlsx");
            _excelapp.Quit();

            #region old
            //_excelapp.SheetsInNewWorkbook = 1;
            //var workbookForContracts = _excelapp.Workbooks.Add(Type.Missing);

            //var outexcelsheets = workbookForContracts.Worksheets;
            ////Получаем ссылку на лист 1
            //var outexcelworksheet = (Excel.Worksheet)outexcelsheets.get_Item(Convert.ToInt32(textBoxSheet.Text));
            ////Выбираем ячейку для вывода A1/

            //var firstTitle = excelworksheet.get_Range("A1", "A1");
            ////Выводим число
            //firstTitle.Value2 = "FirstTitle";

            //var secondTitle = excelworksheet.get_Range("B1", "B1");
            ////Выводим число
            //firstTitle.Value2 = "SecondTitle";

            //var thirdTitle = excelworksheet.get_Range("C1", "C1");
            ////Выводим число
            //firstTitle.Value2 = "ThirdTitle";

            //workbookForContracts.Saved = true;
            //_excelapp.DisplayAlerts = false;
            //workbookForContracts.SaveAs(@"E:\notjob\ExcelTestApp\output1.xlsx",  //object Filename
            //   Excel.XlFileFormat.xlHtml,          //object FileFormat
            //   Type.Missing,                       //object Password 
            //   Type.Missing,                       //object WriteResPassword  
            //   Type.Missing,                       //object ReadOnlyRecommended
            //   Type.Missing,                       //object CreateBackup
            //   Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
            //   Type.Missing,                       //object ConflictResolution
            //   Type.Missing,                       //object AddToMru 
            //   Type.Missing,                       //object TextCodepage
            //   Type.Missing,                       //object TextVisualLayout
            //   Type.Missing);                      //object Local

            //workbookForContracts.Close();

            //_excelapp.Quit();



            //var test = excelworksheet.Cells[2, 2].ToString();

            //var range = excelworksheet.UsedRange.Rows.Count;

            //var port = textBox1.Text;

            //var portList = new List<String>();

            //foreach (var row in excelworksheet.UsedRange.Rows)
            //{

            //}

            //for (int i = 2; i <= range; i++)
            //{
            //    var temp = excelworksheet.get_Range(port + i.ToString(), Type.Missing).Value2;
            //    portList.Add(Convert.ToString(temp));
            //}



            //var good = textBox2.Text;
            //var masa = textBox3.Text;

            //var excelcells = excelworksheet.get_Range(port + "2", Type.Missing);
            //var t1 = Convert.ToString(excelcells.Value2);



            //excelcells = excelworksheet.get_Range("B2", Type.Missing);
            //var t2 = Convert.ToString(excelcells.Value2);

            //excelcells = excelworksheet.get_Range("C2", Type.Missing);
            //var t3 = Convert.ToInt32(excelcells.Value2);

            #endregion

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var application = new Excel.Application {Visible = true};
            //Получаем набор ссылок на объекты Workbook
            var excelappworkbooks = application.Workbooks;
            //Открываем книгу и получаем на нее ссылку
            var excelappworkbook = application.Workbooks.Add();
            //var excelappworkbook = application.Workbooks.Open(@"E:\notjob\ExcelTestApp\aa.xlsx",
            //                   Type.Missing, Type.Missing, Type.Missing,
            // "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
            //  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //  Type.Missing, Type.Missing);
            //Если бы мы открыли несколько книг, то получили ссылку так
            //excelappworkbook=excelappworkbooks[1];
            //Получаем массив ссылок на листы выбранной книги
            var excelsheets = excelappworkbook.Worksheets;
            //Получаем ссылку на лист 1
            var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            //Выбираем ячейку для вывода A1
            var excelcells = excelworksheet.get_Range("A1", "A1");
            //Выводим число
            excelcells.Value2 = 10.5;

            excelappworkbooks = application.Workbooks;
            excelappworkbook = excelappworkbooks[1];
            excelappworkbook.SaveAs(@"E:\notjob\ExcelTestApp\aa.xlsx");
            application.Quit();
        }

        private string GetExcelPath()
        {
            var sourceExcelFileName = String.Empty;

            using (var selectFileDialog = new OpenFileDialog())
            {
                if (selectFileDialog.ShowDialog() == DialogResult.OK)
                {
                    sourceExcelFileName = selectFileDialog.FileName;
                }
            }

            return sourceExcelFileName;
        }
    }
}
