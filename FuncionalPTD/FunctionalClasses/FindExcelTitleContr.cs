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
    /// класс нахождения названия компании в Excel-файле генподрядчика
    /// </summary>
    public class FindExcelTitleContr : FindTitleBehavior
    {
        private int CountingLine { get; set; }
        private int CountingColumn { get; set; }

        private Excel.Application TempImportExcel { get; set; }

        /// <summary>
        /// метод нахождения названия компании в Excel-файле генподрядчика
        /// </summary>
        /// <returns></returns>
        public string FindTitle(Excel.Application TempImportExcel, int index)
        {
            if (this.TempImportExcel == null)
                this.TempImportExcel = TempImportExcel;
            if (CountingLine == 0 || CountingColumn == 0)
            {
                Excel.Range leftTopCell = findLeftTopCell();
                CountingLine = leftTopCell.Row;
                CountingColumn = leftTopCell.Column + 1;

                for (int i = 1; TempImportExcel.Cells[CountingLine + 1, leftTopCell.Column].Text.Trim() != "1"; i++)
                    CountingLine++;
            }

            string Return = TempImportExcel.Cells[CountingLine + index, CountingColumn].Text;
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
