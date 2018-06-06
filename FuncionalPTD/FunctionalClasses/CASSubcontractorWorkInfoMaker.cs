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
    public class CASExcelSubcontrWorkInfoMaker : IWorkInfoMaker
    {
        private CASExcelParserSubcontr parser { get; set; }
            = new CASExcelParserSubcontr();

        public Work MakeInfoWork (Excel.Application TempImportExcel, int index)
        {
            Work result = new Work();

            result.AllocMoney = parser.FindAllocMoney(TempImportExcel, index);
            result.PeriodList = parser.FindPeriodList(TempImportExcel, index);
            result.Title = parser.FindTitle(TempImportExcel, index);
            return result;
        }
    }
}
