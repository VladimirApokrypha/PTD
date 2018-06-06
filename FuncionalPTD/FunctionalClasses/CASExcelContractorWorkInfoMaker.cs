using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DomainPTD.DomainClasses;
using FuncionalPTD.FunctionalInterfaces;
using Excel = Microsoft.Office.Interop.Excel;

namespace FuncionalPTD.FunctionalClasses
{
    public class CASExcelContrWorkInfoMaker : IWorkInfoMaker
    {

        private CASExcelParserContr parser { get; set; }
               = new CASExcelParserContr();


        public Work MakeInfoWork(Excel.Application TempImportExcel,int index )
        {
            Work Result = new Work();
            Result.AllocMoney = parser.FindAllocMoney(TempImportExcel,index);
            Result.Title = parser.FindTitle(TempImportExcel, index);
            Result.PeriodList = parser.FindPeriodList(TempImportExcel, index);
            return Result;
        }
    }
}
