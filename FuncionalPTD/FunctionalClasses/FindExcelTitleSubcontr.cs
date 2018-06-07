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

        private Excel.Application TempImportExcel { get; set; }

        /// <summary>
        /// метод нахождения названия компании в Excel-файле генподрядчика
        /// </summary>
        /// <returns></returns>
        public CASTitle FindTitle(Excel.Application TempImportExcel, int index)
        {
            if (CountingColumn == 0 || CountingLine == 0)
            {
                if (this.TempImportExcel == null)
                    this.TempImportExcel = TempImportExcel;

                Excel.Range leftTopCell = findLeftTopCell();
                CountingLine = leftTopCell.Row + 1;
                CountingColumn = leftTopCell.Column + 2;

                for (int i = 1;
                    TempImportExcel.Cells[leftTopCell.Row + i, leftTopCell.Column].Text.Trim() != "1"; i++)
                    CountingLine++;
            }

            CASTitle result = new CASTitle();
            result.Title = TempImportExcel.Cells[findIndexLine(index), CountingColumn].Text;
            result.Point = false;
            return result;
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
                if (TempImportExcel.Cells[resultIndex, CountingColumn - 1].Value != null)
                    index--;
            }
            return resultIndex;
        }
    }
}
