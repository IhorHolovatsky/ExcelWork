using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.OpenXmlFormats.Spreadsheet;

namespace COFCO.BLL
{
    public static class ContractBL
    {
        /// <summary>
        /// Feels summary in contarct excell
        /// </summary>
        /// <param name="supplierContractsOutputList">list of contracts</param>
        /// <param name="contractsWorksheet">worksheet</param>
        public static void FeelContractsSummary(List<int> supplierContractsOutputList, Excel.Worksheet contractsWorksheet)
        {
            int lastIterationRowNumber = 1;

            foreach (var supplierRowNumber in supplierContractsOutputList)
            {
                var contractsDictionary = new Dictionary<string, double>();

                for (int i = lastIterationRowNumber + 1; i < supplierRowNumber - 1; i++)
                {
                    var contractCell = contractsWorksheet.Cells[i, 9];
                    var contractRange = contractsWorksheet.Range[contractCell, contractCell];
                    var contractValue = contractRange.Value2?.ToString();

                    var quantityCell = contractsWorksheet.Cells[i, 4];
                    var quantityRange = contractsWorksheet.Range[quantityCell, quantityCell];
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
                var outputRange = contractsWorksheet.Range[outputAdress, outputAdress];
                outputRange.Value2 = outputString;

                lastIterationRowNumber = supplierRowNumber;



            }
        }

        public static string GetRowAdressByRange(Excel.Range range)
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
