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
    /// Класс, описывающий файл с работой генподрядчика
    /// </summary>
    [Serializable]
    public class ContrWorkFile : IWorkFile
    {
        private string _path;

        /// <summary>
        /// Путь к файлу
        /// </summary>
        public string Path
        {
            get => _path;
            set
            {
                _path = value;
                if (Path != null)
                    Extension = Path.Split('.').Last();
            }
        }

        public string Extension { get; set; }

        public IWorker worker { get; set; }
            = new Contractor();

        public override string ToString()
        {
            return Path.Split('.')[Path.Split('.').Length - 1];
        }
    }
}
