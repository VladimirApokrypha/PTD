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
            Excel.Application TempImportExcel = new Excel.Application();
            Excel.Workbook TempWoorkBook =
            TempImportExcel.Application.Workbooks.Open(@"C:\Users\Андрей\Desktop\Ik22_Fakt_Vypolnenie_Sogl_Dogovora_Gp_Na_07_11_2017 (2).xlsx");
            Excel.Worksheet TempWorkSheet = TempWoorkBook.Worksheets.get_Item(1);
            TempImportExcel.DisplayAlerts = false;

            CASFileMaker ebaal = new CASFileMaker();
            CASInfoMakerContr suka = new CASInfoMakerContr();
            List<Work> shokal = new List<Work>();

            for (int i = 1; i < 120; i++)
            {
                shokal.Add(suka.MakeInfoWork(TempImportExcel, i));
            }
            TempImportExcel = new Excel.Application();
            TempWoorkBook =
            TempImportExcel.Application.Workbooks.Open(@"C:\Users\Андрей\Desktop\13.07.2017.xlsx");
            TempWorkSheet = TempWoorkBook.Worksheets.get_Item(1);
            TempImportExcel.DisplayAlerts = false;
            List<List<Work>> AllWork = new List<List<Work>>();
            CASInfoMakerSubcontr sub = new CASInfoMakerSubcontr();
            List<Work> subrabota = new List<Work>();
            for (int i = 1; i < 18; i++)
            {
                subrabota.Add(sub.MakeInfoWork(TempImportExcel, i));
            }
            AllWork.Add(subrabota);
            TempWoorkBook.Close(false);
            TempImportExcel.Quit();
            TempImportExcel = null;
            TempWoorkBook = null;
            TempWorkSheet = null;
            GC.Collect();
            TempImportExcel = new Excel.Application();
            TempWoorkBook =
            TempImportExcel.Application.Workbooks.Open(@"C:\Users\Андрей\Desktop\EBAAAAAl.xlsx");
            TempWorkSheet = TempWoorkBook.Worksheets.get_Item(1);
            TempImportExcel.DisplayAlerts = false;
            ebaal.loop(TempImportExcel, shokal,AllWork);
            TempWorkSheet.SaveAs(@"C:\Users\Андрей\Desktop\EBAAAAAl.xlsx");

            

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
