﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FuncionalPTD.FunctionalInterfaces;
using FuncionalPTD.FunctionalInterfaces.Behaviors;
using DomainPTD.DomainClasses;
using Excel = Microsoft.Office.Interop.Excel;

namespace FuncionalPTD.FunctionalClasses
{
    /// <summary>
    /// Класс парсера файлов генподрядчика
    /// </summary>
    public class CASExcelParserSubcontr : ICASParser
    {
        public FindTitleBehavior FindTitleBehavior { get; set; }
        public FindPeriodListBehavior FindPeriodListBehavior { get; set; }
        public FindAllocMoneyBehavior FindAllocMoneyBehavior { get; set; }

        public CASExcelParserSubcontr()
        {
            FindTitleBehavior = new FindExcelTitleSubcontr();
            FindPeriodListBehavior = new FindExcelPeriodListSubcontr();
            FindAllocMoneyBehavior = new FindExcelAllocMoneySubcontr();
        }

        /// <summary>
        /// Поиск названия работы в файле субподрядчика
        /// </summary>
        /// <param name="path"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public string FindTitle(Excel.Application TempImportExcel, int index)
        {
            return FindTitleBehavior.FindTitle(TempImportExcel, index);
        }

        /// <summary>
        /// Поиск сроков выполнения работы в файле субподрядчика
        /// </summary>
        /// <param name="path"></param>
        /// <param name="workTitle"></param>
        /// <returns></returns>
        public List<Period> FindPeriodList(Excel.Application TempImportExcel, int index)
        {
            return FindPeriodListBehavior.FindPeriodList(TempImportExcel, index);
        }

        /// <summary>
        /// Поиск выделенных дна работу денег в файле субподрядчика
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
            for (result = 0; FindTitleBehavior.FindTitle(TempImportExcel, result) != ""; result++) ;
            return result;
        }
    }
}
