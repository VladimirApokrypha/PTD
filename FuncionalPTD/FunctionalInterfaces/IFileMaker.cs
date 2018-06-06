using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DomainPTD.DomainInterfaces;

namespace FuncionalPTD.FunctionalInterfaces
{
    public interface IFileMaker
    {
        IWorkFile MakeFile(List<IWorkFile> workFileList);
    }
}
