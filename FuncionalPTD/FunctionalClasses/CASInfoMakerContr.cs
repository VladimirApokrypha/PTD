using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using FuncionalPTD.FunctionalClasses;
using DomainPTD.DomainClasses;
using FuncionalPTD.FunctionalInterfaces;

namespace FuncionalPTD.FunctionalClasses
{
    public class CASInfoMakerContr : IWorkInfoMaker
    {
        private CASExcelParserContr parser
            = new CASExcelParserContr();

        public Work MakeInfoWork(object[,] array, int index)
        {
            Work result = new Work();
            result.AllocMoney = parser.FindAllocMoney(array, index);
            result.Title = parser.FindTitle(array, index);
            result.PeriodList = parser.FindPeriodList(array, index);

            return result;
        }
    }
}
