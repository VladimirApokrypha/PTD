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
    public delegate IWorker MakeWorkFunction(IWorkFile workFileList);
    public delegate void MakeFileFunction(Contractor contrWork, List<Subcontractor> subcontrWorks, string path);

    public class CASFileMaker //: IFileMaker
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

        public void MakeFile(List<IWorkFile> workFileList, string path)
        {
            List<Subcontractor> allSubcrotractors = new List<Subcontractor>();
            Contractor contractor = new Contractor();

            foreach (var temp in workFileList)
            {
                MakeWorkFunction makeWorkFunction = TypeWorkList[temp.Extension];
                if (temp is SubcontrWorkFile)
                    allSubcrotractors.Add((Subcontractor)makeWorkFunction(temp));
                else
                    contractor = (Contractor)makeWorkFunction(temp);
            }

            string extension = path.Split('.').Last();
            MakeFileFunction makeFileFunction = TypeOutFile[extension];
            makeFileFunction(contractor, allSubcrotractors, path);
        }

        public IWorker ExcelMakeWorkList(IWorkFile workFile)
        {
            Excel.Application TempImportExcel = new Excel.Application(); ;
            Excel.Workbook TempWoorkBook =
            TempImportExcel.Application.Workbooks.Open(workFile.Path);
            Excel.Worksheet TempWorkSheet = TempWoorkBook.Worksheets.get_Item(1);
            TempImportExcel.DisplayAlerts = false;

            IWorker worker = workFile.worker;
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

            worker.WorkList = works;
            return worker;
        }

        public void ExcelLoop(Contractor contrWork, List<Subcontractor> subcontrWorks, string path)
        {
            Excel.Application TempImportExcel = new Excel.Application(); ;
            Excel.Workbook TempWoorkBook =
            TempImportExcel.Application.Workbooks.Add(1);
            Excel.Worksheet TempWorkSheet = TempWoorkBook.Worksheets.get_Item(1);
            TempImportExcel.DisplayAlerts = false;
            TempImportExcel.Visible = true;

            int coutingColumn = 2;
            int coutingLine = 18;

            for (int i = 0; i < subcontrWorks.Count; i++)
                TempImportExcel.Cells[12, 7 + i].Value = subcontrWorks[i].Name;

            for (int i = 0, index = 1; i < contrWork.WorkList.Count; i++)
            {
                TempImportExcel.Cells[i + coutingLine, coutingColumn].Value = contrWork.WorkList[i].Title.Title;
                if (contrWork.WorkList[i].Title.Point == true)
                    TempImportExcel.Cells[i + coutingLine, coutingColumn - 1].Value = index++;

                coutingColumn++;
                decimal sum = 0;

                TempImportExcel.Cells[i + coutingLine, coutingColumn].Value = contrWork.WorkList[i].AllocMoney;

                coutingColumn += 4;
                for (int j = 0; j < subcontrWorks.Count; j++)
                {
                    
                    for (int k = 0; k < subcontrWorks[j].WorkList.Count; k++)
                    {
                        if ((subcontrWorks[j]).WorkList[k].Title.Title.Trim() == contrWork.WorkList[i].Title.Title.Trim())
                            TempImportExcel.Cells[i + coutingLine, coutingColumn + j].Value = subcontrWorks[j].WorkList[k].AllocMoney;
                    }

                    if (TempImportExcel.Cells[i + coutingLine, 7 + j].Value != null)
                        sum += (decimal)TempImportExcel.Cells[i + coutingLine, coutingColumn + j].Value;
                }

                if (sum != 0) TempImportExcel.Cells[i + coutingLine, coutingColumn + subcontrWorks.Count].Value = sum;

                if (TempImportExcel.Cells[i + coutingLine, coutingColumn + subcontrWorks.Count].Value != null && TempImportExcel.Cells[i + coutingLine, 4].Value != null)
                    TempImportExcel.Cells[i + coutingLine, coutingColumn + subcontrWorks.Count + 1].Value = (decimal)TempImportExcel.Cells[i + coutingLine, 4].Value - (decimal)TempImportExcel.Cells[i + coutingLine, coutingColumn + subcontrWorks.Count].Value;

                coutingColumn = 2;
                coutingLine = 18;
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
