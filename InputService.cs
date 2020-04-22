using GenerateRaportProject.Models;
using System;
using System.Collections.Generic;
using System.Text;
using IronXL;
using System.Linq;

namespace GenerateRaportProject.Services
{
    class InputService
    {
        /** ReadRange zwraca dane z zakresu: A2 tj. początkowa komórka z danymi, A1876 końcowa komórka danych w Excel.*/
        public static Range ReadRange(string path)
        {
            WorkBook workbook;
            try
            {
                 workbook = WorkBook.Load(path);
   
            /** Wyjątek dla otwartego pliku */
            }catch(Exception e)
            {
                Console.WriteLine("Source File is Open");
                Console.WriteLine("Please close Source File");
                Console.ReadKey();
                workbook = WorkBook.Load(path);
            }
            WorkSheet sheet = workbook.WorkSheets.First();
            /** Pobieram dane z komorek od A2 do A1876 */
            Range range = sheet["A2:A1876"];
            return range;

        }
        /** ReadDataFromSource - zwraca dane w postaci listy string do fukcji main, jej parametrem jest sciezka do pliku */
        public static List<string> ReadDataFromSource(string path)
        {
            List<string> inputClients = new List<string>();
            Range range = InputService.ReadRange(path);
            foreach (var cell in range)
            {
                inputClients.Add(cell.Value.ToString());
            }
            return inputClients;
        }
    }
}
