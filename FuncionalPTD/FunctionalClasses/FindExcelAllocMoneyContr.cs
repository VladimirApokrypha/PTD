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

        private Excel.Application TempImportExcel { get; set; }

        public decimal FindAllocMoney(Excel.Application TempImportExcel, int index)
        {
            if (this.TempImportExcel == null)
                this.TempImportExcel = TempImportExcel;
            if (CountingLine == 0 || CountingColumn == 0)
            {
                Excel.Range leftTopCell = findLeftTopCell();
                CountingLine = leftTopCell.Row;
                CountingColumn = leftTopCell.Column + 2;

                for (int i = 1; TempImportExcel.Cells[CountingLine + 1, leftTopCell.Column].Text.Trim() != "1"; i++)
                    CountingLine++;
            }
            decimal Return = 0;
            if (TempImportExcel.Cells[CountingLine + index, CountingColumn].Value != null)
                Return = (decimal)TempImportExcel.Cells[CountingLine + index, CountingColumn].Value;
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
    }
}
