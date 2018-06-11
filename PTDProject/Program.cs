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
            manager.AddCASFile(@"C:\Users\Владимир\Desktop\test folder\test project\test file.xlsx");
        }
    }
}
