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
                Visible = false
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
            
            int firstColumn = Convert.ToInt32(column1textBox.Text);
            int secondColumn = Convert.ToInt32(column2textBox.Text);
            int thirdColumn = Convert.ToInt32(column3textBox.Text);


            int b = 0;

            var elementsList = new List<MainInfoModel>();
            foreach (Excel.Range row in excelworksheet.UsedRange.Rows)
            {
                b++;
                if (b == 1) { continue;
                    
                }
                String[] rowData = new String[row.Columns.Count+1];
                for (int i = 1; i <= row.Columns.Count; i++)
                {
                    var v1 = row.Cells[1, i];
                    var v2 = v1?.Value2;
                    var v3 = v2?.ToString();
                    rowData[i] = v3;
                }


                elementsList.Add(new MainInfoModel()
                {
                    FirstString = rowData[firstColumn],
                    SecondString = rowData[secondColumn],
                    FirstNumber = Convert.ToDouble(rowData[thirdColumn])
                });

                if (b == 100)
                {
                    break;
                    
                }

            }

            _excelapp.Workbooks[1].Close();

            CreateTempBookForContracts(elementsList);
           
           // _excelapp.Quit();
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

        private void CreateTempBookForContracts(List<MainInfoModel> list)
        {
            var outputExcelappWorkbook = _excelapp.Workbooks.Add();

            //Получаем массив ссылок на листы выбранной книги
            var outputExcelSheets = outputExcelappWorkbook.Worksheets;

            //Получаем ссылку на лист 1
            var outputExcelWorksheet = (Excel.Worksheet)outputExcelSheets.Item[1];
            
            var firstTitle = outputExcelWorksheet.Range["A1", "A1"];
            firstTitle.Value2 = "FirstTitle";

            var secondTitle = outputExcelWorksheet.Range["B1", "B1"];
            secondTitle.Value2 = "SecondTitle";

            var thirdTitle = outputExcelWorksheet.Range["C1", "C1"];
            thirdTitle.Value2 = "ThirdTitle";

            var fourthTitle = outputExcelWorksheet.Range["D1", "D1"];
            fourthTitle.Value2 = "Contracts";

            int rowIteration = 1;

            foreach (MainInfoModel t in list)
            {
                rowIteration++;
                var cellLiteralA = "A" + rowIteration;
                var cellLiteralB = "B" + rowIteration;
                var cellLiteralC = "C" + rowIteration;
                outputExcelWorksheet.Range[cellLiteralA, cellLiteralA].Value2 = t.FirstString;
                outputExcelWorksheet.Range[cellLiteralB, cellLiteralB].Value2 = t.SecondString;
                outputExcelWorksheet.Range[cellLiteralC, cellLiteralC].Value2 = t.FirstNumber;
            }


            var outexcelappworkbooks = _excelapp.Workbooks;
            var outexcelappworkbook = outexcelappworkbooks[1];

            var outputPath = string.Empty;

            using (var selectFolderDialog = new FolderBrowserDialog())
            {
                if (selectFolderDialog.ShowDialog() == DialogResult.OK)
                {
                    outputPath = selectFolderDialog.SelectedPath;
                }
            }

            outexcelappworkbook.SaveAs(outputPath + "\\tempContracts.xlsx");
            _excelapp.Visible = true;
        }
    }
}
