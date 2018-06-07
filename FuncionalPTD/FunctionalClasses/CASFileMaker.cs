using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FuncionalPTD.FunctionalClasses;
using FuncionalPTD.FunctionalInterfaces;
using DomainPTD.DomainClasses;
using Excel = Microsoft.Office.Interop.Excel;
using DomainPTD.DomainInterfaces;

namespace FuncionalPTD.FunctionalClasses
{
    public delegate List<Work> MakeWorkFunction(IWorkFile workFileList);
    public delegate CASCombine MakeFileFunction(List<Work> contrWork, List<List<Work>> subcontrWorks, string path);

    public class CASFileMaker : IFileMaker
    {
        public CASFileMaker()
        {
            TypeFileList.Add("xlsx", ExcelMakeWorkList);
            TypeOutFile.Add("xlsx", ExcelLoop);
        }

        public Dictionary<string, MakeWorkFunction> TypeFileList { get; set; }
            = new Dictionary<string, MakeWorkFunction>();

        public Dictionary<string, MakeFileFunction> TypeOutFile { get; set; }
            = new Dictionary<string, MakeFileFunction>();

        public IWorkFile MakeFile(List<IWorkFile> workFileList, string path)
        {
            List<List<Work>> allWorks = new List<List<Work>>();
            CASCombine result = new CASCombine();
            
            foreach (var temp in workFileList)
            {
                MakeWorkFunction function = TypeFileList[temp.Extension];
                allWorks.Add(function(temp));
            }

            return result;
        }

        public List<Work> ExcelMakeWorkList(IWorkFile workFile)
        {
            Excel.Application TempImportExcel = new Excel.Application(); ;
            Excel.Workbook TempWoorkBook =
            TempImportExcel.Application.Workbooks.Open(workFile.Path);
            Excel.Worksheet TempWorkSheet = TempWoorkBook.Worksheets.get_Item(1);
            TempImportExcel.DisplayAlerts = false;

            List<Work> works = new List<Work>();

            if (workFile is Subcontractor)
            {
                CASExcelParserSubcontr parser = new CASExcelParserSubcontr();
                CASInfoMakerSubcontr infoMaker = new CASInfoMakerSubcontr();
                for (int index = 0; index < parser.LastIndexInFile(TempImportExcel); index++)
                    works.Add(infoMaker.MakeInfoWork(TempImportExcel, index));
            }
            else
            {
                CASExcelParserContr parser = new CASExcelParserContr();
                CASInfoMakerContr infoMaker = new CASInfoMakerContr();
                for (int index = 0; index < parser.LastIndexInFile(TempImportExcel); index++)
                    works.Add(infoMaker.MakeInfoWork(TempImportExcel, index));
            }

            return works;
        }

        public CASCombine ExcelLoop(List<Work> contrWork, List<List<Work>> subcontrWorks, string path)
        {
            Excel.Application TempImportExcel = new Excel.Application(); ;
            Excel.Workbook TempWoorkBook =
            TempImportExcel.Application.Workbooks.Open(path);
            Excel.Worksheet TempWorkSheet = TempWoorkBook.Worksheets.get_Item(1);
            TempImportExcel.DisplayAlerts = false;

            return new CASCombine();
        }
    }
}
