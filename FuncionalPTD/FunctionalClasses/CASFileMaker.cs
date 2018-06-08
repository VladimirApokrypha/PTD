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

            Excel.Range range;



            range = TempWorkSheet.get_Range("B10", "B12");
            range.EntireColumn.ColumnWidth = 8;
            range.EntireRow.RowHeight = 20;
            ExcelFormat(range);
            range.Value2 = "№ п.п.";

            range = TempWorkSheet.get_Range("C10", "C12");
            range.EntireColumn.ColumnWidth = 36;
            ExcelFormat(range);
            range.Value2 = "Наименование работ";

            range = TempWorkSheet.get_Range("D10", "D12");
            range.EntireColumn.ColumnWidth = 16;
            ExcelFormat(range);
            range.Value2 = "стоимость работ ВСЕГО ГП";

            range = TempWorkSheet.get_Range("E10", "F10");
            range.EntireColumn.ColumnWidth = 17;
            range.EntireRow.RowHeight = 30;
            ExcelFormat(range);
            range.Value2 = "Стоимость работ на торги";

            range = TempWorkSheet.get_Range("E11", "E12");
            range.EntireRow.RowHeight = 15;
            ExcelFormat(range);
            range.Value2 = "Выставлено";

            range = TempWorkSheet.get_Range("F11", "F12");
            ExcelFormat(range);
            range.Value2 = "Маржа(ГП – СП)";

            range = TempWorkSheet.get_Range("A13", "A13");
            range.EntireRow.RowHeight = 35;

            range = TempWorkSheet.get_Range("C14", "C15");
            ExcelFormat(range);
            range.Value2 = "ВСЕГО стоимость, руб.";

            range = TempWorkSheet.get_Range("G10", (object)TempImportExcel.Cells[10, 7 + subcontrWorks.Count]);
            ExcelFormat(range);
            range.Value2 = "Стоимость работ всего СП //по договору";

            range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[11, 7 + subcontrWorks.Count], (object)TempImportExcel.Cells[12, 7 + subcontrWorks.Count]);
            range.EntireColumn.ColumnWidth = 20;
            ExcelFormat(range);
            range.Value2 = "Итог СП";

            range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[10, 7 + subcontrWorks.Count + 1], (object)TempImportExcel.Cells[12, 7 + subcontrWorks.Count + 1]);
            range.EntireColumn.ColumnWidth = 18;
            ExcelFormat(range);
            range.Value2 = "Разница стоимости //по договору";

            int LastColumn = 0;
            int year = contrWork.WorkList[0].PeriodList[0].Date.Year;
            for (int i = 0, h = 11; i < contrWork.WorkList[0].PeriodList.Count; i++, h += 3 + subcontrWorks.Count)
            {
                if (contrWork.WorkList[0].PeriodList[i].Date.Year == year)
                {
                    if (i < 2)
                    {
                        ExcelCapForWork(h, contrWork.WorkList[0].PeriodList[i].Date.ToLongDateString(), "Выполнение работ " + contrWork.WorkList[0].PeriodList[i].Date.ToLongDateString() + " СП", "Разница стоимости за " + contrWork.WorkList[0].PeriodList[i].Date.ToLongDateString(), "Итог " + contrWork.WorkList[0].PeriodList[i].Date.ToLongDateString() + " СП", TempImportExcel, TempWorkSheet, range, subcontrWorks);

                        range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[10, h], (object)TempImportExcel.Cells[12, h]);
                        range.UnMerge();
                        range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[10, h], (object)TempImportExcel.Cells[10, h]);
                        range.EntireColumn.ColumnWidth = 18;
                        ExcelFormat(range);
                        range.Value2 = "Выполнение ГП";

                        range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[11, h], (object)TempImportExcel.Cells[12, h]);
                        range.UnMerge();
                        range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[11, h], (object)TempImportExcel.Cells[12, h]);
                        range.EntireColumn.ColumnWidth = 18;
                        ExcelFormat(range);
                        range.Value2 = null;
                        range.Value2 = contrWork.WorkList[0].PeriodList[i].Date.ToLongDateString();
                    }
                    else
                    {
                        ExcelCapForWork(h, "Накопительно выпонение ГП на " + contrWork.WorkList[0].PeriodList[i].Date.ToLongDateString(), "Накопительно выпонение СП за " + contrWork.WorkList[0].PeriodList[i].Date.ToLongDateString(), "Накопительно разница (гп – сп) на " + contrWork.WorkList[0].PeriodList[i].Date.ToLongDateString(), "Итого накопительно СП", TempImportExcel, TempWorkSheet, range, subcontrWorks);
                    }
                }
                else
                {
                    year++;
                    ExcelCapForWork(h, "Итого ЗАКРЫТО ГП за " + contrWork.WorkList[0].PeriodList[i - 1].Date.Year.ToString(), "Итого закрыто СП за " + contrWork.WorkList[0].PeriodList[i - 1].Date.Year, "Разница стоимости закрытого объема (ГП – СП) за " + contrWork.WorkList[0].PeriodList[i - 1].Date.Year, "Итого закрыто СП за " + contrWork.WorkList[0].PeriodList[i - 1].Date.Year, TempImportExcel, TempWorkSheet, range, subcontrWorks);
                }
                LastColumn = h;
            }

            ExcelCapForWork(LastColumn, "Итого ЗАКРЫТО ГП за " + contrWork.WorkList[0].PeriodList[contrWork.WorkList[0].PeriodList.Count - 1].Date.Year.ToString(), "Итого закрыто СП за " + contrWork.WorkList[0].PeriodList[contrWork.WorkList[0].PeriodList.Count - 1].Date.Year, "Разница стоимости закрытого объема (ГП – СП) за " + contrWork.WorkList[0].PeriodList[contrWork.WorkList[0].PeriodList.Count - 1].Date.Year, "Итого закрыто СП за " + contrWork.WorkList[0].PeriodList[contrWork.WorkList[0].PeriodList.Count - 1].Date.Year, TempImportExcel, TempWorkSheet, range, subcontrWorks);

            LastColumn += 3 + subcontrWorks.Count;
            ExcelCapForWork(LastColumn, "Итого договор минус закрытый объем по ГП ", "Итого договор минус закрытый объем по СП", "Разница стоимости закрытого объема(ГП – СП)", "Итого закрыто СП", TempImportExcel, TempWorkSheet, range, subcontrWorks);

            for (int i = 2; i < LastColumn; i++)
            {
                range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[13, i], (object)TempImportExcel.Cells[13, i]);
                ExcelFormat(range);
                range.Value2 = i - 1;

                range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[14, i], (object)TempImportExcel.Cells[15, i]);
                range.Merge();
                range.Borders.ColorIndex = 1;
                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Borders.Weight = Excel.XlBorderWeight.xlThin;
            }

            int coutingColumn = 3;
            int coutingLine = 18;

            for (int i = 0; i < subcontrWorks.Count; i++)
            {
                Excel.Range temp = TempWorkSheet.Range[TempImportExcel.Cells[11, 7 + i], TempImportExcel.Cells[12, 7 + i]];
                ExcelFormat(temp);
                temp.EntireColumn.ColumnWidth = 20;
                TempImportExcel.Cells[11, 7 + i].Value = subcontrWorks[i].Name;
            }

            for (int i = 0, index = 1; i < contrWork.WorkList.Count; i++)
            {
                TempImportExcel.Cells[i + coutingLine, coutingColumn].Value = contrWork.WorkList[i].Title.Title;
                if (contrWork.WorkList[i].Title.Point == true)
                    TempImportExcel.Cells[i + coutingLine, coutingColumn - 1].Value = index++;

                coutingColumn++;

                if (contrWork.WorkList[i].AllocMoney != 0)
                    TempImportExcel.Cells[i + coutingLine, coutingColumn].Value = contrWork.WorkList[i].AllocMoney;
                else
                    TempImportExcel.Cells[i + coutingLine, coutingColumn].Value = " ";

                decimal sum = 0;

                coutingColumn += 3;
                for (int j = 0; j < subcontrWorks.Count; j++)
                {
                    
                    for (int k = 0; k < subcontrWorks[j].WorkList.Count; k++)
                    {
                        if ((subcontrWorks[j]).WorkList[k].Title.Title == contrWork.WorkList[i].Title.Title
                            && (subcontrWorks[j]).WorkList[k].AllocMoney != 0)
                            TempImportExcel.Cells[i + coutingLine, coutingColumn + j].Value = subcontrWorks[j].WorkList[k].AllocMoney;
                    }

                    if (TempImportExcel.Cells[i + coutingLine, 7 + j].Value != null)
                        sum += (decimal)TempImportExcel.Cells[i + coutingLine, coutingColumn + j].Value;
                }

                if (sum != 0) TempImportExcel.Cells[i + coutingLine, coutingColumn + subcontrWorks.Count].Value = sum;

                if (TempImportExcel.Cells[i + coutingLine, coutingColumn + subcontrWorks.Count].Value != null
                    && TempImportExcel.Cells[i + coutingLine, 4].Value != null)
                    TempImportExcel.Cells[i + coutingLine, coutingColumn + subcontrWorks.Count + 1].Value =
                        (decimal)TempImportExcel.Cells[i + coutingLine, 4].Value
                        - (decimal)TempImportExcel.Cells[i + coutingLine, coutingColumn + subcontrWorks.Count].Value;

                coutingColumn = 3;
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

        private void ExcelFormat(Excel.Range inputrange)
        {
            inputrange.Merge(Type.Missing);
            inputrange.HorizontalAlignment = Excel.Constants.xlCenter;
            inputrange.VerticalAlignment = Excel.Constants.xlCenter;
            inputrange.Borders.ColorIndex = 1;
            inputrange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            inputrange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            inputrange.WrapText = true;
        }

        private void ExcelCapForWork(int column, string period, string naimenovanie, string raznica, string itog, Excel.Application inExcel, Excel.Worksheet inworksheet, Excel.Range inrange, List<Subcontractor> inworks)
        {
            inrange = inworksheet.get_Range((object)inExcel.Cells[10, column], (object)inExcel.Cells[12, column]);
            inrange.EntireColumn.ColumnWidth = 18;
            ExcelFormat(inrange);
            inrange.Value = period;

            inrange = inworksheet.get_Range((object)inExcel.Cells[10, column + 1], (object)inExcel.Cells[10, column + 1 + inworks.Count]);
            ExcelFormat(inrange);
            inrange.Value2 = naimenovanie;

            for (int j = 0; j < inworks.Count; j++)
            {
                inrange = inworksheet.get_Range((object)inExcel.Cells[11, column + 1 + j], (object)inExcel.Cells[12, column + 1 + j]);
                inrange.EntireColumn.ColumnWidth = 20;
                ExcelFormat(inrange);
                inExcel.Cells[11, column + 1 + j].Value2 = inworks[j].Name;
            }

            inrange = inworksheet.get_Range((object)inExcel.Cells[11, column + 1 + inworks.Count], (object)inExcel.Cells[12, column + 1 + inworks.Count]);
            inrange.EntireColumn.ColumnWidth = 20;
            ExcelFormat(inrange);
            inrange.Value2 = itog;

            inrange = inworksheet.get_Range((object)inExcel.Cells[10, column + 1 + inworks.Count + 1], (object)inExcel.Cells[12, column + 1 + inworks.Count + 1]);
            inrange.EntireColumn.ColumnWidth = 18;
            ExcelFormat(inrange);
            inrange.Value2 = raznica;

        }
    }
}
