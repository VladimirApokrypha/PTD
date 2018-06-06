using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DomainPTD.DomainInterfaces;

namespace DomainPTD.DomainClasses
{
    class CASCombine : IWorkFile
    {
        public string Path
        {
            get
            {
                return _path;
            }
            set
            {
                _path = value;
                Extension = _path.Split(new Char[] { '.' }).LastOrDefault();
            }
        }
        string _path;
        public string Extension { get; set; }
        public IWorker worker { get; set; }
    }
}
