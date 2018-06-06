using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DomainPTD.DomainClasses;
using FuncionalPTD.FunctionalInterfaces;
using FuncionalPTD.FunctionalClasses;

namespace FuncionalPTD.FunctionalClasses
{
    public class CASInfoMakerSubcontr : IWorkInfoMaker
    {
        private CASExcelParserSubcontr parser
            = new CASExcelParserSubcontr();

        public Work MakeInfoWork(Excel.Application TempImportExcel, int index)
        {
            Work result = new Work();
            result.AllocMoney = parser.FindAllocMoney(TempImportExcel, index);
            result.Title = parser.FindTitle(TempImportExcel, index);
            result.PeriodList = parser.FindPeriodList(TempImportExcel, index);
            return result;
        }
    }
}
