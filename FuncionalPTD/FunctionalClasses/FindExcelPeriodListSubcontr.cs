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
                "январь", "янв" ,
                "февраль", "фев" ,
                "март", "мар" ,
                "апрель", "апр" ,
                "май", "май" ,
                "июнь", "июн" ,
                "июль", "июль" ,
                "август", "авг",
                "сентябрь", "сен" ,
                "октябрь", "окт" ,
                "ноябрь", "ноя" ,
                "декабрь", "дек"
            };

        private object[,] array { get; set; }

        private int CountingColumn { get; set; }
        private int CountingLine { get; set; }
        private List<DateTime> PeriodsList { get; set; }
            = new List<DateTime>();
        private List<int> ColumnsPeriodsList { get; set; }
            = new List<int>();
        private List<int> DeletedIndexes { get; set; }
            = new List<int>();

        /// <summary>
        /// метод нахождения периодов в Excel файле субподрядчика
        /// </summary>
        /// <param name="path"></param>
        public List<Period> FindPeriodList(object[,] array, int index)
        {
            if (this.array == null)
            {
                this.array = array;
                Cell leftTop = findLeftTopCell();
                leftTop.Row--;
                while (array[leftTop.Row - 1, leftTop.Column] == null)
                    leftTop.Row--;
                for (int i = leftTop.Row; i < array.Length / array.GetLength(1); i++)
                {
                    bool swt = false;
                    for (int j = 1; j <= array.GetLength(1); j++)
                    {
                        DateTime date = findDate(array[i, j]);
                        if (date.Year != 2)
                        {
                            leftTop.Row = i;
                            leftTop.Column = j;
                            swt = true;
                            break;
                        }
                    }
                    if (swt) break;
                }
                DateTime lastDate = new DateTime();
                for (int i = leftTop.Column; i < array.GetLength(1); i++)
                {
                    DateTime newDate = findDate(array[leftTop.Row, i]);
                    if (newDate.Year != 2)
                    {
                        PeriodsList.Add(newDate);
                        ColumnsPeriodsList.Add(i);
                        lastDate = newDate;
                    }
                    else if (!topCell(leftTop.Row, i))
                    {
                        PeriodsList.Add(lastDate);
                        ColumnsPeriodsList.Add(i);
                    }

                }
                SupplementYears(PeriodsList);
            }

            for (int i = 1, j = 0; i < PeriodsList.Count; i++, j++)
            {
                if (PeriodsList[i].Month == PeriodsList[i - 1].Month)
                {
                    DeletedIndexes.Add(i - 1);
                    PeriodsList.RemoveAt(i - 1);
                    i--;
                }
            }

            FindExcelTitleSubcontr finder = new FindExcelTitleSubcontr();
            int line = CellOf(finder.FindTitle(array, index).Title).Row;
            List<decimal> moneys = new List<decimal>();
            decimal sum = 0;
            for (int i = 0; i < ColumnsPeriodsList.Count; i++)
            {
                if (DeletedIndexes.Exists(x => x == i))
                {
                    if (array[line, ColumnsPeriodsList[i]] != null)
                        sum += decimal.Parse(array[line, ColumnsPeriodsList[i]].ToString());
                    DeletedIndexes.Remove(i);
                    ColumnsPeriodsList.RemoveAt(i);
                    i--;
                }
                else
                {
                    if (array[line, ColumnsPeriodsList[i]] != null)
                        moneys.Add(sum + decimal.Parse(array[line, ColumnsPeriodsList[i]].ToString()));
                    else moneys.Add(sum);
                    sum = 0;
                }
            }

            List<Period> result = new List<Period>();
            for (int i = 0; i < PeriodsList.Count; i++)
            {
                Period newPeriod = new Period()
                {
                    Date = PeriodsList[i],
                    Money = moneys[i]
                };
                result.Add(newPeriod);
            }

            return result;
        }

        private DateTime findDate(object period)
        {
            if (period is DateTime)
                return (DateTime)period;
            else if (period is string)
            {
                var month = from temp in monthList
                            where ((string)period).ToLower().Contains(temp)
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

        private Cell findLeftTopCell()
        {
            Cell result;
            int lineIndex = 1, columnIndex = 1;
            for (lineIndex = 1; ; lineIndex++)
            {
                for (columnIndex = 1; columnIndex <= 5; columnIndex++)
                {
                    if (array[lineIndex, columnIndex] != null
                        && array[lineIndex, columnIndex].ToString().Trim() == "1")
                    {
                        result.Row = lineIndex;
                        result.Column = columnIndex;
                        return result;
                    }
                }
            }
        }

        private Cell CellOf(object element)
        {
            Cell result;
            for (int i = 1; i <= array.Length / array.GetLength(1); i++)
            {
                for (int j = 1; j <= array.GetLength(1); j++)
                {
                    if (array[i, j] != null &&
                        array[i, j] == element)
                    {
                        result.Column = j;
                        result.Row = i;
                        return result;
                    }
                }
            }
            result.Column = -1;
            result.Row = -1;
            return result;
        }

        private bool topCell(int i, int j)
        {
            for (int r = j; r > i - 2; r--)
                if (array[r, j] != null)
                    return true;
            return false;
        }
    }
}
