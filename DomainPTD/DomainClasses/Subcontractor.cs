using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DomainPTD.DomainInterfaces;

namespace DomainPTD.DomainClasses
{
    /// <summary>
    /// класс, описывающий данные субподрядчика
    /// </summary>
    [Serializable]
    public class Subcontractor : IWorker
    {
        /// <summary>
        /// название компании генподрядчика
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// список всех работ
        /// </summary>
        public List<Work> WorkList { get; set; }
            = new List<Work>();
    }
}