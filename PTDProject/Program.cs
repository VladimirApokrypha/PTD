using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FuncionalPTD.FunctionalClasses;
using DomainPTD.DomainClasses;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace PTDProject
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application TempImportExcel = new Excel.Application(); ;
            Excel.Workbook TempWoorkBook =
            TempImportExcel.Application.Workbooks.Open(@"C:\Users\Владимир\Desktop\Лист Microsoft Excel (2).xlsx");
            Excel.Worksheet TempWorkSheet = TempWoorkBook.Worksheets.get_Item(1);
            TempImportExcel.DisplayAlerts = false;

            TempImportExcel.Cells[1, 1] = "2";

            Console.ReadKey();
            TempWorkSheet.SaveAs(@"C:\Users\Владимир\Desktop\Лист Microsoft Excel (2).xlsx");
            TempWoorkBook.Close(false);
            TempImportExcel.Quit();
            TempImportExcel = null;
            TempWoorkBook = null;
            TempWorkSheet = null;
            GC.Collect();
        }
    }
}
