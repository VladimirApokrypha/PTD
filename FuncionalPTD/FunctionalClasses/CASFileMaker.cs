﻿using System;
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
            object[,] array = TempWorkSheet.UsedRange.Value;

            IWorker worker = workFile.worker;
            List<Work> works = new List<Work>();

            if (workFile is SubcontrWorkFile)
            {
                CASExcelParserSubcontr parser = new CASExcelParserSubcontr();
                CASInfoMakerSubcontr infoMaker = new CASInfoMakerSubcontr();
                int lastIndexInFile = parser.LastIndexInFile(array);
                for (int index = 1; index <= lastIndexInFile; index++)
                    works.Add(infoMaker.MakeInfoWork(array, index));
            }
            else
            {
                CASExcelParserContr parser = new CASExcelParserContr();
                CASInfoMakerContr infoMaker = new CASInfoMakerContr();
                int lastIndexInFile = parser.LastIndexInFile(array);
                for (int index = 1; index <= lastIndexInFile; index++)
                    works.Add(infoMaker.MakeInfoWork(array, index));
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

            Excel.Range range;
            object[,] array = new object[contrWork.WorkList.Count + 20, contrWork.WorkList[1].PeriodList.Count * (subcontrWorks.Count + 3) * 6];
            int coutingColumn = 2;
            int coutingLine = 17;

            array[13, 2] = "ВСЕГО стоимость, руб.";

            for (int i = 0; i < subcontrWorks.Count; i++)
            {
                Excel.Range temp = TempWorkSheet.Range[TempImportExcel.Cells[11, 7 + i], TempImportExcel.Cells[12, 7 + i]];
                ExcelFormat(temp);
                temp.EntireColumn.ColumnWidth = 20;
                array[11, 7 + i] = subcontrWorks[i].Name;
            }

            for (int i = 0, index = 1; i < contrWork.WorkList.Count; i++)
            {
                array[i + coutingLine, coutingColumn] = contrWork.WorkList[i].Title.Title;
                if (contrWork.WorkList[i].Title.Point == true)
                {
                    array[i + coutingLine, coutingColumn - 1] = index++;
                    Excel.Range temp = TempWorkSheet.Range[TempImportExcel.Cells[i + coutingLine + 1, 2], TempImportExcel.Cells[i + coutingLine + 1, 3]];
                    temp.EntireRow.Font.Bold = true;
                    Excel.Range temp2 = TempWorkSheet.Range[TempImportExcel.Cells[i + coutingLine + 1, 2], TempImportExcel.Cells[i + coutingLine + 1, 2]];
                    temp2.HorizontalAlignment = Excel.Constants.xlCenter;
                }
                else
                {
                    array[i + coutingLine, coutingColumn - 1] = " ";
                    Excel.Range temp = TempWorkSheet.Range[TempImportExcel.Cells[i + coutingLine + 1, 3], TempImportExcel.Cells[i + coutingLine + 1, 3]];
                    temp.Font.Italic = true;
                    temp.HorizontalAlignment = Excel.Constants.xlRight;
                    temp.WrapText = true;
                    temp.RowHeight = 0;
                }

                coutingColumn++;

                if (contrWork.WorkList[i].AllocMoney != 0)
                    array[i + coutingLine, coutingColumn] = contrWork.WorkList[i].AllocMoney;
                else
                    array[i + coutingLine, coutingColumn] = " ";

                decimal sum = 0;
                decimal prevContrSum = 0;
                decimal[] prevSubcontrSum = new decimal[subcontrWorks.Count];

                coutingColumn += 3;
                for (int j = 0; j < subcontrWorks.Count; j++)
                {
                    for (int k = 0; k < subcontrWorks[j].WorkList.Count; k++)
                    {
                        if (subcontrWorks[j].WorkList[k].Title.Title == contrWork.WorkList[i].Title.Title
                            && subcontrWorks[j].WorkList[k].AllocMoney != 0)
                        {
                            array[i + coutingLine, coutingColumn + j] = subcontrWorks[j].WorkList[k].AllocMoney;
                            sum += subcontrWorks[j].WorkList[k].AllocMoney;
                        }
                    }
                }

                if (sum != 0) array[i + coutingLine, coutingColumn + subcontrWorks.Count] = sum;
                array[i + coutingLine, coutingColumn + subcontrWorks.Count + 1] = contrWork.WorkList[i].AllocMoney - sum;

                coutingColumn += subcontrWorks.Count + 2;
                sum = 0;

                for (int j = 0; j < contrWork.WorkList[i].PeriodList.Count; j++)
                {
                    array[i + coutingLine, coutingColumn++] = contrWork.WorkList[i].PeriodList[j].Money.Value;
                    prevContrSum += contrWork.WorkList[i].PeriodList[j].Money.Value;
                    for (int k = 0; k < subcontrWorks.Count; k++)
                    {
                        for (int r = 0; r < subcontrWorks[k].WorkList[0].PeriodList.Count; r++)
                        {

                            Work work = subcontrWorks[k].WorkList.Find
                                (x => x.Title.Title == contrWork.WorkList[i].Title.Title);
                            Period period = null;
                            if (work != null)
                            {
                                period = work.PeriodList.Find(x => x.Date == contrWork.WorkList[i].PeriodList[j].Date);
                                if (period != null)
                                {
                                    array[i + coutingLine, coutingColumn + k] = period.Money.Value;
                                    prevSubcontrSum[k] += period.Money.Value;
                                    sum += period.Money.Value;
                                    break;
                                }
                            }
                        }
                    }

                    coutingColumn += subcontrWorks.Count;

                    if (sum != 0) array[i + coutingLine, coutingColumn] = sum;
                    array[i + coutingLine, coutingColumn + 1] = contrWork.WorkList[i].PeriodList[j].Money - sum;
                    coutingColumn += 2;
                    sum = 0;
                    if (j != 0)
                    {
                        if (prevContrSum != 0)
                            array[i + coutingLine, coutingColumn] = prevContrSum;
                        coutingColumn++;
                        for (int r = 0; r < subcontrWorks.Count; r++, coutingColumn++)
                        {
                            if (prevSubcontrSum[r] != 0)
                            {
                                array[i + coutingLine, coutingColumn] = prevSubcontrSum[r];
                                sum += prevSubcontrSum[r];
                            }
                        }

                        if (sum != 0) array[i + coutingLine, coutingColumn] = sum;
                        array[i + coutingLine, coutingColumn + 1] = prevContrSum - sum;
                        coutingColumn += 2;
                    }
                    sum = 0;
                }
                
                decimal resultContrSum = contrWork.WorkList[i].AllocMoney - prevContrSum;
                array[i + coutingLine, coutingColumn++] = resultContrSum;
                for (int j = 0; j < subcontrWorks.Count; j++, coutingColumn++)
                {
                    for (int k = 0; k < subcontrWorks[j].WorkList.Count; k++)
                    {
                        Work work = subcontrWorks[j].WorkList.Find
                                (x => x.Title.Title == contrWork.WorkList[i].Title.Title);
                        if (work != null)
                        {
                            array[i + coutingLine, coutingColumn] = work.AllocMoney - prevSubcontrSum[j];
                            sum += work.AllocMoney - prevSubcontrSum[j];
                        }
                    }
                }

                if (sum != 0) array[i + coutingLine, coutingColumn] = sum;
                array[i + coutingLine, coutingColumn + 1] = resultContrSum - sum;
                coutingColumn += 2;

                prevContrSum = 0;
                coutingColumn = 2;
                coutingLine = 17;
            }


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
            range.Value2 = "Cтоимость работ ВСЕГО ГП";

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
            range.Value2 = "Стоимость работ всего СП";

            range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[11, 7 + subcontrWorks.Count], (object)TempImportExcel.Cells[12, 7 + subcontrWorks.Count]);
            range.EntireColumn.ColumnWidth = 20;
            ExcelFormat(range);
            range.Value2 = "Итог СП";

            range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[10, 7 + subcontrWorks.Count + 1], (object)TempImportExcel.Cells[12, 7 + subcontrWorks.Count + 1]);
            range.EntireColumn.ColumnWidth = 18;
            ExcelFormat(range);
            range.Value2 = "Разница стоимости";

            for (int i = 0; i < subcontrWorks.Count; i++)
            {
                range = TempWorkSheet.Range[TempWorkSheet.Cells[11, 7 + i], TempWorkSheet.Cells[12, 7 + i]];
                range.Value = subcontrWorks[i].Name;
                ExcelFormat(range);
            }

            int LastColumn = 0;
            int year = contrWork.WorkList[0].PeriodList[0].Date.Year;
            for (int i = 0, h = 7 + subcontrWorks.Count + 2, z = 0; i < contrWork.WorkList[0].PeriodList.Count; i++, h += 3 + subcontrWorks.Count)
            {
                if (contrWork.WorkList[0].PeriodList[i].Date.Year == year)
                {
                    if (z == 0)
                    {
                        ExcelCapForWork(h, contrWork.WorkList[0].PeriodList[i].Date.ToString("y"), "Выполнение работ " + contrWork.WorkList[0].PeriodList[i].Date.ToString("y") + " СП", "Разница стоимости на " + contrWork.WorkList[0].PeriodList[i].Date.ToString("y"), "Итог СП", TempImportExcel, TempWorkSheet, range, subcontrWorks);

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
                        range.Value2 = contrWork.WorkList[0].PeriodList[i].Date.ToString("y");
                        z++;
                    }
                    else
                    {
                        ExcelCapForWork(h, contrWork.WorkList[0].PeriodList[i].Date.ToString("y"), "Выполнение работ " + contrWork.WorkList[0].PeriodList[i].Date.ToString("y") + " СП", "Разница стоимости на " + contrWork.WorkList[0].PeriodList[i].Date.ToString("y"), "Итог СП", TempImportExcel, TempWorkSheet, range, subcontrWorks);

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
                        range.Value2 = contrWork.WorkList[0].PeriodList[i].Date.ToString("y");

                        h += 3 + subcontrWorks.Count;

                        ExcelCapForWork(h, "Накопительно выпонение ГП на " + contrWork.WorkList[0].PeriodList[i].Date.ToString("y"), "Накопительно выпонение СП на " + contrWork.WorkList[0].PeriodList[i].Date.ToString("y"), "Накопительно разница (гп – сп) на " + contrWork.WorkList[0].PeriodList[i].Date.ToString("y"), "Итого накопительно СП", TempImportExcel, TempWorkSheet, range, subcontrWorks);
                    }
                }
                else
                {
                    h = h - 3 - subcontrWorks.Count;
                    year++;
                    ExcelCapForWork(h, "Итого ЗАКРЫТО ГП за " + contrWork.WorkList[0].PeriodList[i - 1].Date.Year.ToString(), "Итого закрыто СП за " + contrWork.WorkList[0].PeriodList[i - 1].Date.Year, "Разница стоимости закрытого объема (ГП – СП) за " + contrWork.WorkList[0].PeriodList[i - 1].Date.Year, "Итого закрыто ", TempImportExcel, TempWorkSheet, range, subcontrWorks);
                    i--;
                }
                LastColumn = h;
            }

            ExcelCapForWork(LastColumn, "Итого ЗАКРЫТО ГП за " + contrWork.WorkList[0].PeriodList[contrWork.WorkList[0].PeriodList.Count - 1].Date.Year.ToString(), "Итого закрыто СП за " + contrWork.WorkList[0].PeriodList[contrWork.WorkList[0].PeriodList.Count - 1].Date.Year, "Разница стоимости закрытого объема (ГП – СП) за " + contrWork.WorkList[0].PeriodList[contrWork.WorkList[0].PeriodList.Count - 1].Date.Year, "Итого закрыто СП", TempImportExcel, TempWorkSheet, range, subcontrWorks);

            LastColumn += 3 + subcontrWorks.Count;
            ExcelCapForWork(LastColumn, "ИТОГО договор минус закрытый объем по ГП ", "ИТОГО договор минус закрытый объем по СП", "ИТОГО разница стоимости закрытого объема(ГП – СП)", "ИТОГО закрыто СП", TempImportExcel, TempWorkSheet, range, subcontrWorks);

            LastColumn += 3 + subcontrWorks.Count;
            for (int i = 2; i < LastColumn; i++)
            {
                range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[13, i], (object)TempImportExcel.Cells[13, i]);
                ExcelFormat(range);
                range.Value2 = i - 1;

                range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[14, i], (object)TempImportExcel.Cells[15, i]);
                range.Merge();
            }
            range = TempWorkSheet.Range[(object)TempImportExcel.Cells[10, 2], (object)TempImportExcel.Cells[17 + contrWork.WorkList.Count, LastColumn - 1]];
            range.Borders.ColorIndex = 1;
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.Weight = Excel.XlBorderWeight.xlThin;


            coutingColumn = 3;
            coutingLine = 17;

            decimal allSum = 0;
            
            for (int i = coutingColumn; i < contrWork.WorkList[1].PeriodList.Count * (subcontrWorks.Count + 3) * 6; i++)
            {
                for (int j = coutingLine; j < contrWork.WorkList.Count + 17; j++)
                {
                    if (array[j, i] is decimal) allSum += (decimal)array[j, i];
                }
                if (allSum != 0)
                    array[13, i] = allSum;
                allSum = 0;
            }

            for (int i = 0; i < array.Length / array.GetLength(1); i++)
            {
                for (int j = 0; j < array.GetLength(1); j++)
                {
                    if (array[i, j] != null && array[i, j].ToString() == "0")
                        array[i, j] = "";
                }
            }

            object[,] newArray = new object[contrWork.WorkList.Count + 20, contrWork.WorkList[1].PeriodList.Count * (subcontrWorks.Count + 3) * 6];

            for (int i = 13, newI = 0; i < contrWork.WorkList.Count + 20; i++, newI++)
            {
                for (int j = 1, newJ = 0; j < contrWork.WorkList[1].PeriodList.Count * (subcontrWorks.Count + 3) * 6; j++, newJ++)
                {
                    newArray[newI, newJ] = array[i, j];
                }
            }


            Excel.Range rangee = TempWorkSheet.Range
                [TempWorkSheet.Cells[14, 2],
                TempWorkSheet.Cells[contrWork.WorkList.Count + 20, contrWork.WorkList[1].PeriodList.Count * (subcontrWorks.Count + 3) * 6]];

            rangee.Value = newArray;

            range = TempWorkSheet.get_Range((object)TempImportExcel.Cells[16, 2], (object)TempImportExcel.Cells[17, LastColumn - 1]);
            range.Interior.ColorIndex = 6;
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
            inputrange.Merge();
            inputrange.HorizontalAlignment = Excel.Constants.xlCenter;
            inputrange.VerticalAlignment = Excel.Constants.xlCenter;
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
