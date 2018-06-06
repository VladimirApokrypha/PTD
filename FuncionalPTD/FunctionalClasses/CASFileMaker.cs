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
    public delegate List<Work> MakeFunction(IWorkFile workFileList);

    public class CASFileMaker : IFileMaker
    {
        public CASFileMaker()
        {
            TypeFileList.Add("xlsx", ExcelMakeWorkList);
        }

        public Dictionary<string, MakeFunction> TypeFileList { get; set; }
            = new Dictionary<string, MakeFunction>();

        public IWorkFile MakeFile(List<IWorkFile> workFileList)
        {
            List<List<Work>> allWorks = new List<List<Work>>();
            foreach (var temp in workFileList)
            {
                MakeFunction function = TypeFileList[temp.Extension];
                allWorks.Add(function(temp));
            }

            CASCombine result = new CASCombine();
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
    }
}
