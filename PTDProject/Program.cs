using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FuncionalPTD.FunctionalClasses;
using DomainPTD.DomainClasses;
using Excel = Microsoft.Office.Interop.Excel;

namespace PTDProject
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application TempImportExcel = new Excel.Application(); ;
            Excel.Workbook TempWoorkBook =
            TempImportExcel.Application.Workbooks.Open(@"C:\Users\Андрей\Desktop\Ik22_Fakt_Vypolnenie_Sogl_Dogovora_Gp_Na_07_11_2017 (2).xlsx");
            Excel.Worksheet TempWorkSheet = TempWoorkBook.Worksheets.get_Item(1);
            TempImportExcel.DisplayAlerts = false;
            
            Console.ReadKey();
            TempWoorkBook.Close(false);
            TempImportExcel.Quit();
            TempImportExcel = null;
            TempWoorkBook = null;
            TempWorkSheet = null;
            GC.Collect();
        }
    }
}
