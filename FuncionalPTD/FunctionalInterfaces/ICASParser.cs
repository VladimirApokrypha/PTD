using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FuncionalPTD.FunctionalInterfaces.Behaviors;

namespace FuncionalPTD.FunctionalInterfaces
{
    /// <summary>
    /// интерфейс описывающий функционал парсера файлов подрядчиков и субподрядчиков
    /// </summary>
    public interface ICASParser
    {
        FindTitleBehavior FindTitleBehavior { get; set; }
        FindPeriodListBehavior FindPeriodListBehavior { get; set; }
        FindAllocMoneyBehavior FindAllocMoneyBehavior { get; set; }
    }
}
