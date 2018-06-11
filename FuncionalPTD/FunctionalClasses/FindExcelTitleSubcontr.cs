using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FuncionalPTD.FunctionalInterfaces.Behaviors;
using Excel = Microsoft.Office.Interop.Excel;
using DomainPTD.DomainClasses;

namespace FuncionalPTD.FunctionalClasses
{
    /// <summary>
    /// класс нахождения названия компании в Excel-файле субподрядчика
    /// </summary>
    public class FindExcelTitleSubcontr : FindTitleBehavior
    {
        private int CountingLine { get; set; }
        private int CountingColumn { get; set; }

        private object[,] array { get; set; }

        /// <summary>
        /// метод нахождения названия компании в Excel-файле генподрядчика
        /// </summary>
        /// <returns></returns>
        public CASTitle FindTitle(object[,] array, int index)
        {
            if (CountingColumn == 0 || CountingLine == 0)
            {
                if (this.array == null)
                    this.array = array;

                Cell leftTopCell = findLeftTopCell();
                CountingLine = leftTopCell.Row + 1;
                CountingColumn = leftTopCell.Column + 2;

                for (int i = 1;
                    array[leftTopCell.Row + i, leftTopCell.Column]?.ToString().Trim() != "1"; i++)
                    CountingLine++;
            }

            CASTitle result = new CASTitle();
            result.Title = (string)array[findIndexLine(index), CountingColumn];
            result.Point = false;
            return result;
        }

        //private Excel.Range findLeftTopCell()
        //{
        //    int lineIndex = 1, columnIndex = 1;
        //    for (lineIndex = 1; ; lineIndex++)
        //    {
        //        for (columnIndex = 1; columnIndex <= 5; columnIndex++)
        //        {
        //            if (TempImportExcel.Cells[lineIndex, columnIndex].Text.Trim() == "1")
        //                return TempImportExcel.Cells[lineIndex, columnIndex];
        //        }
        //    }
        //}

        private Cell findLeftTopCell()
        {
            Cell result;
            int lineIndex = 1, columnIndex = 1;
            for (lineIndex = 1; ; lineIndex++)
            {
                for (columnIndex = 1; columnIndex <= 5; columnIndex++)
                {
                    if (array[lineIndex, columnIndex]?.ToString().Trim() == "1")
                    {
                        result.Row = lineIndex;
                        result.Column = columnIndex;
                        return result;
                    }
                }
            }
        }

        private int findIndexLine(int index)
        {
            int temp = index;
            int resultIndex = CountingLine - 1;
            while (index != 0)
            {
                resultIndex++;
                try
                {
                    if (array[resultIndex, CountingColumn - 1] != null)
                        index--;
                }
                catch(IndexOutOfRangeException ex)
                {
                    return findIndexLine(temp - 1);
                }
            }
            return resultIndex;
        }
    }
}
