using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomainPTD.DomainClasses
{
    [Serializable]
    public class Label
    {
        public string Cell { get; set; }
        public string Title { get; set; }
        public string Info { get; set; }
        public string FilePath { get; set; }
    }
}