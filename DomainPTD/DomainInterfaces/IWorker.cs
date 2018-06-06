using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DomainPTD.DomainClasses;

namespace DomainPTD.DomainInterfaces
{
    /// <summary>
    /// интерфейс исполнителя работы
    /// </summary>
    public interface IWorker
    {
        /// <summary>
        /// название или имя исполнителя
        /// </summary>
        string Name { get; set; }
        List<Work> WorkList { get; set; }
    }
}
