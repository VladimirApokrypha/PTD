using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DomainPTD.DomainInterfaces;

namespace DomainPTD.DomainClasses
{
    /// <summary>
    /// Класс, описывающий файл с работой субподрядчика
    /// </summary>
    [Serializable]
    public class SubcontrWorkFile : IWorkFile
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
            = new Subcontractor();

        public override string ToString()
        {
            return Path.Split('.')[Path.Split('.').Length - 1];
        }
    }
}
