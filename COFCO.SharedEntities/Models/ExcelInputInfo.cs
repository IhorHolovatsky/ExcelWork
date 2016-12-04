using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace COFCO.SharedEntities.Models
{
    public class ExcelInputInfo
    {
        #region File Paths
        public string InputFilePath { get; set; }

        public string OutputTempFolderPath { get; set; }

        public string TempExcelFilePath { get; set; }

        public string OutputTemplateFolderPath { get; set; }
        #endregion

        #region Column Indexes
        public int Port { get; set; }
        public int Supplier { get; set; }
        public int Product { get; set; }
        public int Quantity { get; set; }
        public int Date { get; set; }
        public int VehicleNumber { get; set; }
        public int TTNNumber { get; set; }
        public int Contract { get; set; }
        #endregion

        public int SheetNumber { get; set; }
        public int StartRowNumber { get; set; }
    }
}
