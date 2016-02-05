using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace Analysis
{
    public class MyTable
    {
        private static MyTable uniqueInstance;
        private List<MyRow> list;

        private MyTable()
        {
            list = new List<MyRow>();
        }

        public static MyTable getInstance()
        {
            if (uniqueInstance == null)
                uniqueInstance = new MyTable();

            return uniqueInstance;
        }
        
        public void clearList()
        {
            if (list.Count > 0)
                list.Clear();
        }

        public DataTable ToDataTable()
        {
            return createTable(list);
        }

        public DataTable ToDataTable(string[] numbers)
        {
            var myRows = list.Where(item => item.isEqual(numbers)).ToList();

            return createTable(myRows);
        }

        private DataTable createTable(List<MyRow> myRows)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("артикул");
            dt.Columns.Add("параметр");
            dt.Columns.Add("описание");
            
            foreach (MyRow myRow in myRows)
                dt.Rows.Add(myRow.Row);

            return dt;
        }

        internal void Add(MyRow myRow)
        {
            list.Add(myRow);
        }
    }
}
