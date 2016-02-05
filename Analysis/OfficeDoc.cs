using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Analysis
{
    public class ExcelDoc : OfficeDoc, IDisposable
    {
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkBook;
        private Excel.Worksheet xlSh;

        public ExcelDoc(string name)
            : base(name)
        {
            Init();
            openWorkBook();
        }

        public ExcelDoc()
        {
            Init();
            createWorkBook();
        }

        private void Init()
        {
            xlApp = new Excel.Application();
        }

        private void openWorkBook()
        {
            xlWorkBook = xlApp.Workbooks.Open(name, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlSh = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        }

        private void createWorkBook()
        {
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlSh = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        }

        public void setValue(int rowIndex, int columnIndex, string value)
        {
            xlSh.Cells[rowIndex, columnIndex] = value;
        }

        public object getValue(string rowCell, string columnCell)
        {
            return xlSh.get_Range(rowCell, columnCell).Value2;
        }

        public object[] getRow(int rowIndex)
        {
            return new object[62] {
                            xlSh.get_Range("A" + rowIndex.ToString(), "A" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("B" + rowIndex.ToString(), "B" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("C" + rowIndex.ToString(), "C" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("D" + rowIndex.ToString(), "D" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("E" + rowIndex.ToString(), "E" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("F" + rowIndex.ToString(), "F" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("G" + rowIndex.ToString(), "G" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("H" + rowIndex.ToString(), "H" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("I" + rowIndex.ToString(), "I" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("J" + rowIndex.ToString(), "J" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("K" + rowIndex.ToString(), "K" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("L" + rowIndex.ToString(), "L" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("M" + rowIndex.ToString(), "M" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("N" + rowIndex.ToString(), "N" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("O" + rowIndex.ToString(), "O" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("P" + rowIndex.ToString(), "P" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("Q" + rowIndex.ToString(), "Q" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("R" + rowIndex.ToString(), "R" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("S" + rowIndex.ToString(), "S" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("T" + rowIndex.ToString(), "T" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("U" + rowIndex.ToString(), "U" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("V" + rowIndex.ToString(), "V" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("W" + rowIndex.ToString(), "W" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("X" + rowIndex.ToString(), "X" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("Y" + rowIndex.ToString(), "Y" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("Z" + rowIndex.ToString(), "Z" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AA" + rowIndex.ToString(), "AA" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AB" + rowIndex.ToString(), "AB" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AC" + rowIndex.ToString(), "AC" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AD" + rowIndex.ToString(), "AD" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AE" + rowIndex.ToString(), "AE" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AF" + rowIndex.ToString(), "AF" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AG" + rowIndex.ToString(), "AG" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AH" + rowIndex.ToString(), "AH" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AI" + rowIndex.ToString(), "AI" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AJ" + rowIndex.ToString(), "AJ" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AK" + rowIndex.ToString(), "AK" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AL" + rowIndex.ToString(), "AL" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AM" + rowIndex.ToString(), "AM" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AN" + rowIndex.ToString(), "AN" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AO" + rowIndex.ToString(), "AO" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AP" + rowIndex.ToString(), "AP" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AQ" + rowIndex.ToString(), "AQ" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AR" + rowIndex.ToString(), "AR" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AS" + rowIndex.ToString(), "AS" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AT" + rowIndex.ToString(), "AT" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AU" + rowIndex.ToString(), "AU" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AV" + rowIndex.ToString(), "AV" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AW" + rowIndex.ToString(), "AW" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AX" + rowIndex.ToString(), "AX" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AY" + rowIndex.ToString(), "AY" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("AZ" + rowIndex.ToString(), "AZ" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("BA" + rowIndex.ToString(), "BA" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("BB" + rowIndex.ToString(), "BB" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("BC" + rowIndex.ToString(), "BC" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("BD" + rowIndex.ToString(), "BD" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("BE" + rowIndex.ToString(), "BE" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("BF" + rowIndex.ToString(), "BF" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("BG" + rowIndex.ToString(), "BG" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("BH" + rowIndex.ToString(), "BH" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("BI" + rowIndex.ToString(), "BI" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("BJ" + rowIndex.ToString(), "BJ" + rowIndex.ToString()).Value2
                        };
        }

        public int getCount(string columnName, int rowIndex)
        {
            while(xlSh.get_Range(columnName + rowIndex.ToString(), columnName + rowIndex.ToString()).Value2 != null)
                rowIndex++;

            return rowIndex;
        }

        public void Show()
        {
            xlApp.Visible = true;
        }

        public void Dispose()
        {
            xlApp.Quit();

            releaseObject(xlSh);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
    }

    public abstract class OfficeDoc
    {
        protected string name;

        protected OfficeDoc()
        {
            this.name = "";
        }

        protected OfficeDoc(string name)
        {
            this.name = name;
        }

        protected void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
