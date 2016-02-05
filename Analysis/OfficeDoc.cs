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

        private string _prevValue;

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
            if (xlSh.get_Range("A" + rowIndex.ToString(), "A" + rowIndex.ToString()).Value2 != null)
                _prevValue = xlSh.get_Range("A" + rowIndex.ToString(), "A" + rowIndex.ToString()).Value2.ToString();

            return new object[] {
                            _prevValue,
                            xlSh.get_Range("B" + rowIndex.ToString(), "B" + rowIndex.ToString()).Value2,
                            xlSh.get_Range("C" + rowIndex.ToString(), "C" + rowIndex.ToString()).Value2
                        };
        }

        public int getCount(string columnName, string columnName2, int index)
        {
            while ((xlSh.get_Range(columnName + index.ToString(), columnName + index.ToString()).Value2 != null) || (xlSh.get_Range(columnName + (index + 1).ToString(), columnName + (index + 1).ToString()).Value2 != null)
                || (xlSh.get_Range(columnName2 + index.ToString(), columnName2 + index.ToString()).Value2 != null) || (xlSh.get_Range(columnName2 + (index + 1).ToString(), columnName2 + (index + 1).ToString()).Value2 != null))
            {
                index++;
            }

            return index;
        }

        public void Show()
        {
            xlApp.Visible = true;
        }

        public void Dispose()
        {
            object misValue = System.Reflection.Missing.Value;

            xlApp.DisplayAlerts = false;
            xlApp.EnableEvents = false;

            xlWorkBook.Close(false, misValue, misValue);

            xlApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

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
