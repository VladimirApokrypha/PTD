using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FuncionalPTD.FunctionalInterfaces.Behaviors;
using Excel = Microsoft.Office.Interop.Excel;

namespace FuncionalPTD.FunctionalClasses
{
    /// <summary>
    /// класс нахождения выделенных на работу денег генподрядчика
    /// </summary>
    public class FindExcelAllocMoneyContr : FindAllocMoneyBehavior
    {
        private int CountingLine { get; set; }
        private int CountingColumn { get; set; }

        private object[,] array { get; set; }

        public decimal FindAllocMoney(object[,] array, int index)
        {
            if (this.array == null)
                this.array = array;
            if (CountingLine == 0 || CountingColumn == 0)
            {
                Cell leftTopCell = findLeftTopCell();
                CountingLine = leftTopCell.Row;
                CountingColumn = leftTopCell.Column + 2;

                for (; ; CountingLine++)
                {
                    if (array[CountingLine + 1, leftTopCell.Column] != null
                        && array[CountingLine + 1, leftTopCell.Column].ToString().Trim() == "1")
                        break;
                }
            }
            decimal Return = 0;
            if (array[CountingLine + index, CountingColumn] != null)
                Return = decimal.Parse(array[CountingLine + index, CountingColumn].ToString());
            return Return;
        }

        private Cell findLeftTopCell()
        {
            Cell result;
            int lineIndex = 1, columnIndex = 1;
            for (lineIndex = 1; ; lineIndex++)
            {
                for (columnIndex = 1; columnIndex <= 5; columnIndex++)
                {
                    if (array[lineIndex, columnIndex] != null &&
                        array[lineIndex, columnIndex].ToString().Trim() == "1")
                    {
                        result.Row = lineIndex;
                        result.Column = columnIndex;
                        return result;
                    }
                }
            }
        }
    }
    public struct Cell
    {
        public int Row;
        public int Column;
    }
}
