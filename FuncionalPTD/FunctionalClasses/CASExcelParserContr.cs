using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FuncionalPTD.FunctionalInterfaces;
using FuncionalPTD.FunctionalInterfaces.Behaviors;
using DomainPTD.DomainClasses;
using DomainPTD.DomainInterfaces;
using Excel = Microsoft.Office.Interop.Excel;

namespace FuncionalPTD.FunctionalClasses
{
    /// <summary>
    /// Класс парсера файлов генподрядчика
    /// </summary>
    public class CASExcelParserContr : ICASParser
    {
        public FindTitleBehavior FindTitleBehavior { get; set; }
        public FindPeriodListBehavior FindPeriodListBehavior { get; set; }
        public FindAllocMoneyBehavior FindAllocMoneyBehavior { get; set; }

        public CASExcelParserContr()
        {
            FindTitleBehavior = new FindExcelTitleContr();
            FindPeriodListBehavior = new FindExcelPeriodListContr();
            FindAllocMoneyBehavior = new FindExcelAllocMoneyContr();
        }

        /// <summary>
        /// Поиск названия работы в файле генподрядчика
        /// </summary>
        /// <param name="path"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public CASTitle FindTitle(Excel.Application TempImportExcel, int index)
        {
            return FindTitleBehavior.FindTitle(TempImportExcel, index);
        }

        /// <summary>
        /// Поиск сроков выполнения работы в файле генподрядчика
        /// </summary>
        /// <param name="path"></param>
        /// <param name="workTitle"></param>
        /// <returns></returns>
        public List<Period> FindPeriodList(Excel.Application TempImportExcel, int index)
        {
            return FindPeriodListBehavior.FindPeriodList(TempImportExcel, index);
        }

        /// <summary>
        /// Поиск выделенных дна работу денег в файле генподрядчика
        /// </summary>
        /// <param name="path"></param>
        /// <param name="work"></param>
        /// <returns></returns>
        public decimal FindAllocMoney(Excel.Application TempImportExcel, int index)
        {
            return FindAllocMoneyBehavior.FindAllocMoney(TempImportExcel, index);
        }

        public int LastIndexInFile(Excel.Application TempImportExcel)
        {
            int result;
            for (result = 1; FindTitleBehavior.FindTitle(TempImportExcel, result + 1).Title != ""; result++) ;
            return result;
        }
    }
}
