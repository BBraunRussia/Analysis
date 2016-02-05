using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace Analysis
{
    internal class MyRow
    {
        /*
        private string numberDeclaration;
        private string numberGoods;
        private string custom;
        private DateTime date;
        private string direction;
        private string mode;
        private string senderName;
        private string senderINN;
        private string senderAddress;
        private string recipientName;
        private string recipientINN;
        private string recipientAddress;
        private string declarantName;
        private string declarantINN;
        private string declatantAddress;
        private string count;
        private string numberSeats;
        private string declarationType;
        private string countryWhereCode;
        private string countryTradingCode;
        private string countryDeclarationCode;
        private string countrySenderCode;
        private string countryDestinationCode;
        private string conditionsDelivery;
        private string deliveryPoint;
        private string numberCustomsAtBorder;
        private string manufacturerName;
        private string trademark1;
        private string codeTNVED;
        private string description;
        private string packing;
        private string quantityGoods;
        private string unitMeasurement;
        */
        private string find1;
        private string find2;
        public object[] Row;
        
        public MyRow(DataRow row)
        {
            find1 = row.ItemArray[28].ToString().ToLower().Trim();
            find2 = row.ItemArray[29].ToString().ToLower().Trim();
            loadData(row);
        }

        private void loadData(DataRow row)
        {
            Row = new object[62]
             {
                row.ItemArray[0],
                row.ItemArray[1],
                row.ItemArray[2],
                row.ItemArray[3],
                row.ItemArray[4],
                row.ItemArray[5],
                row.ItemArray[6],
                row.ItemArray[7],
                row.ItemArray[8],
                row.ItemArray[9],
                row.ItemArray[10],
                row.ItemArray[11],
                row.ItemArray[12],
                row.ItemArray[13],
                row.ItemArray[14],
                row.ItemArray[15],
                row.ItemArray[16],
                row.ItemArray[17],
                row.ItemArray[18],
                row.ItemArray[19],
                row.ItemArray[20],
                row.ItemArray[21],
                row.ItemArray[22],
                row.ItemArray[23],
                row.ItemArray[24],
                row.ItemArray[25],
                row.ItemArray[26],
                row.ItemArray[27],
                row.ItemArray[28],
                row.ItemArray[29],
                row.ItemArray[30],
                row.ItemArray[31],
                row.ItemArray[32],
                row.ItemArray[33],
                row.ItemArray[34],
                row.ItemArray[35],
                row.ItemArray[36],
                row.ItemArray[37],
                row.ItemArray[38],
                row.ItemArray[39],
                row.ItemArray[40],
                row.ItemArray[41],
                row.ItemArray[42],
                row.ItemArray[43],
                row.ItemArray[44],
                row.ItemArray[45],
                row.ItemArray[46],
                row.ItemArray[47],
                row.ItemArray[48],
                row.ItemArray[49],
                row.ItemArray[50],
                row.ItemArray[51],
                row.ItemArray[52],
                row.ItemArray[53],
                row.ItemArray[54],
                row.ItemArray[55],
                row.ItemArray[56],
                row.ItemArray[57],
                row.ItemArray[58],
                row.ItemArray[59],
                row.ItemArray[60],
                row.ItemArray[61]
                };
        }

        internal bool isEqual(string[] wordsForSearch)
        {
            foreach (string word in wordsForSearch)
                if ((find1.IndexOf(word.ToLower().Trim()) > -1) || (find2.IndexOf(word.ToLower().Trim()) > -1))
                    return true;

            return false;
        }
    }
}
