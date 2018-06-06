using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FuncionalPTD.FunctionalInterfaces.Behaviors;
using Excel = Microsoft.Office.Interop.Excel;
using DomainPTD.DomainClasses;
using System.Text.RegularExpressions;

namespace FuncionalPTD.FunctionalClasses
{
    /// <summary>
    /// класс нахождения периодов в Excel файле субподрядчика
    /// </summary>
    public class FindExcelPeriodListSubcontr : FindPeriodListBehavior
    {

        private List<string> monthList { set; get; }
            = new List<string>()
            {
                "январь", "янв",
                "февраль", "фев",
                "март", "мар",
                "апрель", "апр",
                "май", "май",
                "июнь", "июн",
                "июль", "июль",
                "август", "авг",
                "сентябрь", "сен",
                "октябрь", "окт",
                "ноябрь", "ноя",
                "декабрь", "дек"
            };

        private Excel.Application TempImportExcel { get; set; }

        private int CountingColumn { get; set; }
        private int CountingLine { get; set; }
        private List<DateTime> PeriodsList { get; set; }
        private List<int> ColumnsPeriodsList { get; set; }
        private List<int> DeletedIndexes { get; set; }
            = new List<int>();

        /// <summary>
        /// метод нахождения периодов в Excel файле субподрядчика
        /// </summary>
        /// <param name="path"></param>
        public List<Period> FindPeriodList(Excel.Application TempImportExcel, int index)
        {
            if (PeriodsList == null)
            {
                if (this.TempImportExcel == null)
                    this.TempImportExcel = TempImportExcel;
                Excel.Range startRange = findLeftTopCell();
                CountingLine = startRange.Row;
                CountingColumn = startRange.Column;

                int CountingPeriodLine = CountingLine - 1;
                int CountingMoneyLine = CountingPeriodLine + 2;

                while (TempImportExcel.Cells[CountingMoneyLine + 1, CountingColumn].Text != "1")
                    CountingMoneyLine++;

                CreatePeriodList();

                for (int i = 1; i < PeriodsList.Count; i++)
                {
                    if (PeriodsList[i].Year == PeriodsList[i - 1].Year
                        && PeriodsList[i].Month == PeriodsList[i - 1].Month)
                    {
                        PeriodsList.RemoveAt(i - 1);
                        DeletedIndexes.Add(i - 1);
                        i--;
                    }
                }
            }

            List<decimal?> moneyList = new List<decimal?>();

            int newLineIndex = findIndexLine(index);

            for (int i = ColumnsPeriodsList.First(); i <= ColumnsPeriodsList.Last(); i++)
            {
                if (ColumnsPeriodsList.Contains(i))
                {
                    if (TempImportExcel.Cells[newLineIndex, i].Value != null)
                        moneyList.Add((decimal)TempImportExcel.Cells[newLineIndex, i].Value);
                    else
                        moneyList.Add(0);
                }
            }

            List<int> tempList = new List<int>();
            tempList.AddRange(DeletedIndexes);
            for (int i = 1; i < moneyList.Count; i++)
            {
                if (tempList.Contains(i - 1))
                {
                    moneyList[i] += moneyList[i - 1];
                    moneyList.RemoveAt(i - 1);
                    tempList.Remove(i - 1);
                    i--;
                }
            }

            List<Period> result = new List<Period>();
            for (int i = 0; i < moneyList.Count; i++)
            {
                result.Add(new Period { Date = PeriodsList[i], Money = moneyList[i] });
            }

            return result;
        }

        private void CreatePeriodList()
        {
            PeriodsList = new List<DateTime>();
            ColumnsPeriodsList = new List<int>();
            for (int i = 0; findCell(CountingLine - 1, CountingColumn + i).Value != null; i++)
            {
                DateTime newDate = findDate(findCell(CountingLine - 1, CountingColumn + i));
                if (newDate.Year != 2)
                {
                    PeriodsList.Add(newDate);
                    ColumnsPeriodsList.Add(CountingColumn + i);
                }
            }
            SupplementYears(PeriodsList);
        }

        private Excel.Range findCell(int row, int column) =>
            TempImportExcel.Cells[TempImportExcel.Cells[row, column].MergeArea.Row,
                TempImportExcel.Cells[row, column].MergeArea.Column];

        private DateTime findDate(Excel.Range period)
        {
            if (period.Value.GetType() == typeof(DateTime))
                return period.Value;
            else if (period.Value.GetType() == typeof(string))
            {
                var month = from temp in monthList
                            where period.Text.ToLower().Contains(temp)
                            select (monthList.IndexOf(temp) + 1) / 2
                            + ((monthList.IndexOf(temp) + 1) % 2 == 0 ? 0 : 1);
                if (month.Count() != 0)
                {
                    return new DateTime(1, month.Last(), 1);
                }
            }
            return new DateTime(2, 1, 1);
        }

        private void SupplementYears(List<DateTime> periods)
        {
            int firstYearInList = 1;
            foreach (DateTime temp in periods)
            {
                if (temp.Year != 1)
                {
                    firstYearInList = temp.Year;
                    for (int i = periods.IndexOf(temp); i != 0; i--)
                        if (periods[i].Month == 12 && i + 1 < periods.Count - 1 && periods[i + 1].Month != 12)
                            firstYearInList--;
                    break;
                }
            }

            bool IsOneDecember = true;
            for (int i = 0; i < periods.Count; i++)
            {
                if (periods[i].Year == 1)
                    periods[i] = new DateTime(firstYearInList, periods[i].Month, 1);
                if (periods[i].Month == 12 && IsOneDecember && i + 1 < periods.Count - 1 && periods[i + 1].Month != 12)
                {
                    firstYearInList++;
                    IsOneDecember = false;
                }
                else if (periods[i].Month != 12) IsOneDecember = true;
            }
        }

        private int findIndexLine(int index)
        {
            int resultIndex = CountingLine + 1;
            for (int i = CountingLine + 1; index != 0; i++)
            {
                if (TempImportExcel.Cells[i, CountingColumn].Value != null)
                {
                    resultIndex = i;
                    index--;
                }
            }
            return resultIndex;
        }

        private Excel.Range findLeftTopCell()
        {
            int lineIndex = 1, columnIndex = 1;
            for (lineIndex = 1; ; lineIndex++)
            {
                for (columnIndex = 1; columnIndex <= 5; columnIndex++)
                {
                    if (TempImportExcel.Cells[lineIndex, columnIndex].Text.Trim() == "1")
                        return TempImportExcel.Cells[lineIndex, columnIndex];
                }
            }
        }
    }
}
