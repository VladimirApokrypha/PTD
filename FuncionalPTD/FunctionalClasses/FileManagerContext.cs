using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;
using DomainPTD.DomainClasses;

namespace FuncionalPTD.FunctionalClasses
{
    public class FileManagerContext : DbContext
    {
        public FileManagerContext() : base("FileManagerConnection") {  }

        public DbSet<ContrWorkFile> ContrWorkFile { set; get; }
        //public DbSet<SubcontrWorkFile> SubcontrWorkFiles { get; set; }
    }
}
