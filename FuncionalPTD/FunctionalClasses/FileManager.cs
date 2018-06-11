using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FuncionalPTD.FunctionalInterfaces.Behaviors;
using DomainPTD.DomainClasses;
using DomainPTD.DomainInterfaces;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;


namespace FuncionalPTD.FunctionalClasses
{
    /// <summary>
    /// класс, описывающий работу менеджера файлов
    /// </summary>
    [Serializable]
    public class FileManager
    {
        public FileManager()
        {
            Serializer serializer = new Serializer();
            FileManager manager = (FileManager)serializer.DeserializeObject(SerializeFile);
            CurrentReportPath = manager.CurrentReportPath;
            ReportPath = manager.ReportPath;
            CurrentResourcesPath = manager.CurrentResourcesPath;
            Contractor = manager.Contractor;
            Subcontractors = manager.Subcontractors;
        }

        private const string ContractorFolderName = "Генподрядчик";
        private const string SubcontractorFolderName = "Субподрядчик";
        private const string ResourcesFolderName = "Ресурсы";
        private const string SerializeFile = "ContractorInfo.dat";

        /// <summary>
        /// путь текущего корневого каталога
        /// </summary>
        public string CurrentReportPath { get; set; } 
        /// <summary>
        /// путь текущего проекта
        /// </summary>
        public string ReportPath { get; set; }
        /// <summary>
        /// путь к данным о текущем проекте
        /// </summary>
        public string CurrentResourcesPath { get; set; }
        /// <summary>
        /// поведение добавление файла соотношения генподрядчика и субподрядчиков
        /// </summary>
        public AddCASFileBehavior AddCASFileBehavior { get; set; }
        /// <summary>
        /// Генподрядчик в проекте
        /// </summary>
        public ContrWorkFile Contractor { get; set; } 
            = new ContrWorkFile();
        /// <summary>
        /// коллекция субподрядчиков в проекте
        /// </summary>
        public ObservableCollection<SubcontrWorkFile> Subcontractors { get; set; }
            = new ObservableCollection<SubcontrWorkFile>();

        /// <summary>
        /// метод создает корневой каталог по указанному пути
        /// </summary>
        /// <param name="path"></param>
        public void CreateGeneralFolder(string path, string name)
        {
            CurrentReportPath = Path.Combine(path, name);
            ReportPath = null;
            CurrentResourcesPath = null;
            Contractor = new ContrWorkFile();
            Subcontractors = new ObservableCollection<SubcontrWorkFile>();

            Directory.CreateDirectory(CurrentReportPath);
        }

        /// <summary>
        /// метод открывает новый корневой каталог
        /// </summary>
        /// <param name="path"></param>
        public void OpenGeneralFolder(string path)
        {
            CurrentReportPath = path;
            ReportPath = null;
            CurrentResourcesPath = null;
            Contractor = new ContrWorkFile();
            Subcontractors = new ObservableCollection<SubcontrWorkFile>();
        }

        public void OpenLastProject()
        {
            Serializer serializer = new Serializer();
            FileManager manager = (FileManager)serializer.DeserializeObject(SerializeFile);
            CurrentReportPath = manager.CurrentReportPath;
            ReportPath = manager.ReportPath;
            CurrentResourcesPath = manager.CurrentResourcesPath;
            Contractor = manager.Contractor;
            Subcontractors = manager.Subcontractors;
        }

        /// <summary>
        /// метод выстраивает древо каталогов при создании проекта
        /// </summary>
        /// <param name="path"></param>
        public void CreateProject(string name)
        {
            ReportPath = Path.Combine(CurrentReportPath, name);
            Directory.CreateDirectory(ReportPath);
            Directory.CreateDirectory(Path.Combine(ReportPath, ContractorFolderName));
            Directory.CreateDirectory(Path.Combine(ReportPath, SubcontractorFolderName));
            Directory.CreateDirectory(Path.Combine(ReportPath, ResourcesFolderName));
        }

        public void OpenProject(string name)
        {
            ReportPath = Path.Combine(CurrentReportPath, name);

            DirectoryInfo contrInfo = new DirectoryInfo(Path.Combine(ReportPath, ContractorFolderName));
            FileInfo[] contr = contrInfo.GetFiles();
            Contractor.Path = Path.Combine(ReportPath, ContractorFolderName, contr[0].Name);

            DirectoryInfo subcontrInfo = new DirectoryInfo(Path.Combine(ReportPath, SubcontractorFolderName));
            foreach (FileInfo temp in subcontrInfo.GetFiles())
            {
                Subcontractors.Add(new SubcontrWorkFile()
                { Path = Path.Combine(ReportPath, SubcontractorFolderName, temp.Name) });
            }
        }

        /// <summary>
        /// метод добавляет нового генподрядчика по указанному пути или создает в случае его отсутствия
        /// </summary>
        /// <param name="path"></param>
        public void AddContractor(string path, string name)
        {
            if (Contractor.Path != null)
            {
                File.Delete(Contractor.Path);
            }
            Contractor.Path = Path.Combine(ReportPath, ContractorFolderName, findFullName(path));
            Contractor.worker.Name = name;
            File.Copy(path, Contractor.Path);
        }

        /// <summary>
        /// метод добавляет нового субподрядчика
        /// </summary>
        /// <param name="path"></param>
        public void AddSubcontractor (string path, string name)
        {
            SubcontrWorkFile newSubcontractor = new SubcontrWorkFile();
            newSubcontractor.Path = Path.Combine(ReportPath, SubcontractorFolderName, findFullName(path));
            newSubcontractor.worker.Name = name;
            foreach (SubcontrWorkFile temp in Subcontractors)
            {
                if (temp.Path == newSubcontractor.Path)
                {
                    File.Delete(temp.Path);
                }
            }
            File.Copy(path, newSubcontractor.Path);
            Subcontractors.Add(newSubcontractor);
        }

        /// <summary>
        /// метод дабавления файла соотношения генподрядчика и субподрядчика
        /// </summary>
        /// <param name="path"></param>
        public void AddCASFile(string path)
        {
            List<IWorkFile> allFiles = new List<IWorkFile>();
            allFiles.Add(Contractor);
            allFiles.AddRange(Subcontractors.ToList());
            CASFileMaker maker = new CASFileMaker();
            maker.MakeFile(allFiles, path);
        }

        public bool IsContractorFileExist()
        {
            if (Contractor != null)
                return false;
            else
                return true;
        }

        private string findFullName(string path)
        {
            string result = path.Split('\\', '/')[path.Split('\\', '/').Length - 1];
            return result;
        }

        public void Serialize()
        {
            Serializer serializer = new Serializer();
            serializer.SerializeObject(SerializeFile, this);
        }
    }
}
