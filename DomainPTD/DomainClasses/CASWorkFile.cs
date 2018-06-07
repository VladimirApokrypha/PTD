using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DomainPTD.DomainInterfaces;

namespace DomainPTD.DomainClasses
{
    /// <summary>
    /// класс описывающий файл соотношения генподрядчика и субподрядчиков
    /// </summary>
    public class CASWorkFile : IWorkFile
    {
        /// <summary>
        /// путь к файлу
        /// </summary>
        public string Path { get; set; }
        /// <summary>
        /// список всех работ
        /// </summary>
        public List<Work> WorkList { get; set; }
            = new List<Work>();
        public string Extension { get; set; }
    }
}
