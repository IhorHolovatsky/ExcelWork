using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using COFCO.SharedEntities.Constants;
using COFCO.SharedEntities.Models;
using NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;

namespace COFCO.BLL
{
    public class ExcelService
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelInputInfo"></param>
        public void CreateTempExcelFile(ExcelInputInfo excelInputInfo)
        {
            if (string.IsNullOrWhiteSpace(excelInputInfo.InputFilePath)
                || string.IsNullOrWhiteSpace(excelInputInfo.OutputTempFolderPath))
            {
                throw new Exception("Input file path or Output temp folder path is empty!");
            }

            var inputFileExtension = Path.GetExtension(excelInputInfo.InputFilePath);
            var isXlsx = string.Equals(inputFileExtension, FileContants.XLSX);

            if (isXlsx)
            {
                #region Logic for XLSX

                XSSFWorkbook hssfwb;

                try
                {
                    using (var file = new FileStream(excelInputInfo.InputFilePath, FileMode.Open, FileAccess.Read))
                    {
                        hssfwb = new XSSFWorkbook(file);
                    }
                }
                catch (OfficeXmlFileException e)
                {
                    throw new Exception("Invalid excel extension");
                }


                var sheet = hssfwb.GetSheetAt(excelInputInfo.SheetNumber);

                for (var rowIndex = excelInputInfo.StartRowNumber; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    var row = sheet.GetRow(rowIndex);
                    
                }
                #endregion
            }
            else
            {

            }
        }
    }
}
