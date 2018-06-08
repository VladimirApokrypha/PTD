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
            FileManager manager = new FileManager();
            //manager.CreateGeneralFolder(@"C:\Users\Владимир\Desktop", "test folder");
            //manager.CreateProject("test project");
            //manager.AddContractor(@"C:\Users\Владимир\Desktop\ИК22 факт выполнение согл договора ГП на 07.11.2017.xlsx");
            //manager.AddSubcontractor(@"C:\Users\Владимир\Desktop\выполнение согл договора Универсал на 13.07.2017.xlsx");
            //manager.AddSubcontractor(@"C:\Users\Владимир\Desktop\выполнение согл договора Универсал на 13.07.2017 — копия.xlsx");
            //manager.Serialize();
            manager.AddCASFile(@"C:\Users\Владимир\Desktop\test folder\test project\test file.xlsx");
            Console.ReadKey();
        }
    }
}
