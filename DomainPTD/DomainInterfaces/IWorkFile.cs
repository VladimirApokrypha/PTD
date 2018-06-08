using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomainPTD.DomainInterfaces
{
    /// <summary>
    /// интерфейс, описывающий файл с работой
    /// </summary>
    public interface IWorkFile
    {
        /// <summary>
        /// путь к файлу
        /// </summary>
        string Path { get; set; }
        string Extension { get; set; }
        IWorker worker { get; set; }
    }
}
