using IronXL;
using System;
using System.Collections.Generic;
using System.Text;

namespace GenerateRaportProject.Services
{
    class OutputService
    {
        /**SaveDataToCSV - zapisuje dane do pliku Raport.xlsx */
        public static void SaveDataToXLSX(List<string> outputClientList) {

            WorkBook xlsxWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
            xlsxWorkbook.Metadata.Author = "Kasper";
            WorkSheet xlsSheet = xlsxWorkbook.CreateWorkSheet("Raport");
            // pentla wpisujaca dane do xlsSheet
            for (int i = 0; i< outputClientList.Count; i++)
            {
                // ustalenie adresu komórki
                string field = "A"+(i+1).ToString();
                xlsSheet[field].Value = outputClientList[i];
            }
            xlsxWorkbook.SaveAs("Raport.xlsx");
        }
    }
}
