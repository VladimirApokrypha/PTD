using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DomainPTD.DomainClasses;
using Excel=Microsoft.Office.Interop.Excel;

namespace FuncionalPTD.FunctionalInterfaces
{
    /// <summary>
    /// интерфейс описывающий функционал составителя информации о работе
    /// </summary>
    public interface IWorkInfoMaker
    {
        /// <summary>
        /// метод возвращает информацию о работе из файла
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        Work MakeInfoWork(object[,] array, int index);
    }
}
