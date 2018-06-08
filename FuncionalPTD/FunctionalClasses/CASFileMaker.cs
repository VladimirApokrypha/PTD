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
    public delegate void MakeFileFunction(List<Work> contrWork, List<List<Work>> subcontrWorks, string path);

    public class CASFileMaker : IFileMaker
    {
        public CASFileMaker()
        {
            TypeWorkList.Add("xlsx", ExcelMakeWorkList);
            TypeOutFile.Add("xlsx", ExcelLoop);
        }

        public Dictionary<string, MakeWorkFunction> TypeWorkList { get; set; }
            = new Dictionary<string, MakeWorkFunction>();

        public Dictionary<string, MakeFileFunction> TypeOutFile { get; set; }
            = new Dictionary<string, MakeFileFunction>();

        public IWorkFile MakeFile(List<IWorkFile> workFileList, string path)
        {
            List<List<Work>> allWorks = new List<List<Work>>();
            List<Work> contrWorks = new List<Work>();
            CASCombine result = new CASCombine();
            result.Path = path;

            foreach (var temp in workFileList)
            {
                MakeWorkFunction makeWorkFunction = TypeWorkList[temp.Extension];
                if (temp is SubcontrWorkFile)
                    allWorks.Add(makeWorkFunction(temp));
                else
                    contrWorks = makeWorkFunction(temp);
            }

            MakeFileFunction makeFileFunction = TypeOutFile[result.Extension];
            makeFileFunction(contrWorks, allWorks, result.Path);

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
            
            if (workFile is SubcontrWorkFile)
            {
                CASExcelParserSubcontr parser = new CASExcelParserSubcontr();
                CASInfoMakerSubcontr infoMaker = new CASInfoMakerSubcontr();
                int lastIndexInFile = parser.LastIndexInFile(TempImportExcel);
                for (int index = 1; index <= lastIndexInFile; index++)
                    works.Add(infoMaker.MakeInfoWork(TempImportExcel, index));
            }
            else
            {
                CASExcelParserContr parser = new CASExcelParserContr();
                CASInfoMakerContr infoMaker = new CASInfoMakerContr();
                int lastIndexInFile = parser.LastIndexInFile(TempImportExcel);
                for (int index = 1; index <= lastIndexInFile; index++)
                    works.Add(infoMaker.MakeInfoWork(TempImportExcel, index));
            }

            TempWoorkBook.Close(false);
            TempImportExcel.Quit();
            TempImportExcel = null;
            TempWoorkBook = null;
            TempWorkSheet = null;
            GC.Collect();

            return works;
        }

        public void ExcelLoop(List<Work> contrWork, List<List<Work>> subcontrWorks, string path)
        {
            Excel.Application TempImportExcel = new Excel.Application(); ;
            Excel.Workbook TempWoorkBook =
            TempImportExcel.Application.Workbooks.Add(1);
            Excel.Worksheet TempWorkSheet = TempWoorkBook.Worksheets.get_Item(1);
            TempImportExcel.DisplayAlerts = false;
            TempImportExcel.Visible = true;

            for (int i = 0; i < contrWork.Count; i++)
            {
                TempImportExcel.Cells[i + 18, 2].Value = contrWork[i].Title.Title;
                if (contrWork[i].Title.Point == true)
                    TempImportExcel.Cells[i + 18, 1].Value = i + 1;

                TempImportExcel.Cells[i + 18, 3].Value = contrWork[i].AllocMoney;
                for (int j = 0; j < subcontrWorks.Count; j++)
                {
                    for (int k = 0; k < subcontrWorks[j].Count; k++)
                    {
                        if ((subcontrWorks[j])[k].Title.Title.Trim() == contrWork[i].Title.Title.Trim())
                        {
                            TempImportExcel.Cells[i + 18, 5 + j].Value = (subcontrWorks[j])[k].AllocMoney;
                        }
                    }
                }
            }

            TempWorkSheet.SaveAs(path);
            TempWoorkBook.Close(false);
            TempImportExcel.Quit();
            TempImportExcel = null;
            TempWoorkBook = null;
            TempWorkSheet = null;
            GC.Collect();
        }
    }
}
