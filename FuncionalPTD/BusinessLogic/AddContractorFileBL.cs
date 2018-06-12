using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GalaSoft.MvvmLight.Command;
using FuncionalPTD.FunctionalClasses;

namespace FuncionalPTD.BusinessLogic
{
    public class AddContractorFileBL
    {
        public AddContractorFileBL()
        {
            manager = new FileManager();
            AddContractorCommand = new RelayCommand(AddContractor, () => Path != "");
        }

        private FileManager manager;

        public string Path { get; set; } = "Enter path to your contractor file";
        public string Title { get; set; } = "Enter name of your contractor";

        public void AddContractor()
        {
            manager.AddContractor(Path, Title);
        }

        public RelayCommand AddContractorCommand { get; set; }
    }
}