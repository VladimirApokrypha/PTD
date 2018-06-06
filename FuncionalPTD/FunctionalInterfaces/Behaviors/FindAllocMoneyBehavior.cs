using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace FuncionalPTD.FunctionalInterfaces.Behaviors
{
    /// <summary>
    /// интерфейс поведения нахождения выделенных денег 
    /// </summary>
    public interface FindAllocMoneyBehavior
    {
        /// <summary>
        /// метод нахождения выделенных на работу денег в файле 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="workTitle"></param>
        /// <returns></returns>
        decimal FindAllocMoney(Excel.Application TempImportExcel, int index);
    }
}
