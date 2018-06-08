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

        private Excel.Application TempImportExcel { get; set; }

        public decimal FindAllocMoney(Excel.Application TempImportExcel, int index)
        {
            if (this.TempImportExcel == null)
                this.TempImportExcel = TempImportExcel;
            if (CountingLine == 0 || CountingColumn == 0)
            {
                Excel.Range leftTopCell = findLeftTopCell();
                CountingLine = leftTopCell.Row;
                CountingColumn = leftTopCell.Column + 3;

                for (int i = 1; TempImportExcel.Cells[CountingLine + 1, leftTopCell.Column].Text.Trim() != "1"; i++)
                    CountingLine++;
            }

            decimal Return = 0;
            if (TempImportExcel.Cells[findIndexLine(index), CountingColumn].Value != null)
                Return = (decimal)TempImportExcel.Cells[findIndexLine(index), CountingColumn].Value;
            return Return;
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

        private int findIndexLine(int index)
        {
            int resultIndex = CountingLine - 1;
            while (index != 0)
            {
                resultIndex++;
                if (TempImportExcel.Cells[resultIndex, CountingColumn - 2].Value != null)
                    index--;
            }
            return resultIndex;
        }
    }
}
