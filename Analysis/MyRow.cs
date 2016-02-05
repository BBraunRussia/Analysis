using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace Analysis
{
    internal class MyRow
    {
        private string find;
        private object[] _row;

        public MyRow(object[] row)
        {
            find = row[0].ToString().ToLower().Trim();
            _row = row;
        }

        public object[] Row { get { return _row; } }

        internal bool isEqual(string[] wordsForSearch)
        {
            foreach (string word in wordsForSearch)
                if ((find.IndexOf(word.ToLower().Trim()) > -1) || (find.IndexOf(word.ToLower().Trim()) > -1))
                    return true;

            return false;
        }
    }
}
