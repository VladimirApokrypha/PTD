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
    /// класс нахождения названия компании в Excel-файле генподрядчика
    /// </summary>
    public class FindExcelTitleContr : FindTitleBehavior
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
            if (this.array == null)
                this.array = array;
            if (CountingLine == 0 || CountingColumn == 0)
            {
                Cell leftTopCell = findLeftTopCell();
                CountingLine = leftTopCell.Row;
                CountingColumn = leftTopCell.Column + 1;

                for (;;CountingLine++)
                {
                    if (array[CountingLine + 1, leftTopCell.Column] != null
                        && array[CountingLine + 1, leftTopCell.Column].ToString().Trim() == "1")
                        break;
                }
            }

            CASTitle result = new CASTitle();
            if (array[CountingLine + index, CountingColumn] != null)
                result.Title = array[CountingLine + index, CountingColumn].ToString();
            else
                result.Title = "";
            if (array[CountingLine + index, CountingColumn - 1] == null)
                result.Point = false;
            else result.Point = true;
            return result;
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
    }
}
