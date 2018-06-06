using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DomainPTD.DomainInterfaces;
using System.ComponentModel.DataAnnotations;

namespace DomainPTD.DomainClasses
{
    /// <summary>
    /// класс, описывающий данные работы
    /// </summary>
    public class Work
    {
        /// <summary>
        /// наименование работы
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// выделенные на работу деньги
        /// </summary>
        public decimal AllocMoney { get; set; }
        /// <summary>
        /// коллекция, содержащая список исполнителей работы
        /// </summary>
        public List<Period> PeriodList { get; set; }
            = new List<Period>();
    }
}
