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
    /// класс нахождения выделенных на работу денег субподрядчика
    /// </summary>
    public class FindExcelAllocMoneySubcontr : FindAllocMoneyBehavior
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
                CountingColumn = leftTopCell.Column + 3;

                for (int i = 1; array[CountingLine + 1, leftTopCell.Column]?.ToString().Trim() != "1"; i++)
                    CountingLine++;
            }

            decimal Return = 0;
            if (array[findIndexLine(index), CountingColumn] != null)
                Return = decimal.Parse(array[findIndexLine(index), CountingColumn]?.ToString());
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
                    if ((array[lineIndex, columnIndex])?.ToString().Trim() == "1")
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
            int resultIndex = CountingLine - 1;
            while (index != 0)
            {
                resultIndex++;
                if (array[resultIndex, CountingColumn - 2] != null)
                    index--;
            }
            return resultIndex;
        }
    }
}
