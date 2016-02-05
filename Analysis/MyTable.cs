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

        public void loadData(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                MyRow myRow = new MyRow(row);
                Add(myRow);
            }
        }

        private void Add(MyRow myRow)
        {
            list.Add(myRow);
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

        public DataTable ToDataTable(string[] trademarks)
        {
            var myRows = from row in list
                         where row.isEqual(trademarks)
                         select row;

            return createTable(myRows.ToList());            
        }

        private DataTable createTable(List<MyRow> myRows)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Номер декларации");
            dt.Columns.Add("№ товара");
            dt.Columns.Add("Таможня");
            dt.Columns.Add("Дата");
            dt.Columns.Add("Направление");
            dt.Columns.Add("Режим");
            dt.Columns.Add("Наименование отправителя");
            dt.Columns.Add("ИНН отправителя");
            dt.Columns.Add("Адрес отправителя");
            dt.Columns.Add("Наименование получателя");
            dt.Columns.Add("ИНН получателя");
            dt.Columns.Add("Адрес получателя");
            dt.Columns.Add("Наименование декларанта");
            dt.Columns.Add("ИНН декларанта");
            dt.Columns.Add("Адрес декларанта");
            dt.Columns.Add("Наименований");
            dt.Columns.Add("Количество мест");
            dt.Columns.Add("Вид декларации");
            dt.Columns.Add("Код страны нахождения");
            dt.Columns.Add("Код торгующей страны");
            dt.Columns.Add("Код страны декл.");
            dt.Columns.Add("Код страны отпр.");
            dt.Columns.Add("Код страны назнач.");
            dt.Columns.Add("Условие поставки");
            dt.Columns.Add("Пункт поставки");
            dt.Columns.Add("Код таможни на границе");

            dt.Columns.Add("Наименование изготовителя");
            dt.Columns.Add("Товарный знак");
            dt.Columns.Add("Код ТН ВЭД");
            dt.Columns.Add("Описание и характеристика товара");
            dt.Columns.Add("Упаковка");
            dt.Columns.Add("Количество товара");
            dt.Columns.Add("Единица измерения");

            dt.Columns.Add("Количество товара1");
            dt.Columns.Add("Единица измерения1");

            dt.Columns.Add("Код единицы измерения");

            dt.Columns.Add("Количество товара2");
            dt.Columns.Add("Единица измерения2");

            dt.Columns.Add("Код единицы измерения1");

            dt.Columns.Add("Код страны происх.");
            dt.Columns.Add("Преференции");
            dt.Columns.Add("КТС");
            dt.Columns.Add("Метод");
            dt.Columns.Add("Платёж (дол)");
            dt.Columns.Add("Платёж (руб)");
            dt.Columns.Add("Фамилия");
            dt.Columns.Add("Телефон");
            dt.Columns.Add("Тип КТС");
            dt.Columns.Add("Должность");
            dt.Columns.Add("Вес нетто (кг)");
            dt.Columns.Add("Вес брутто (кг)");
            dt.Columns.Add("Код валюты тамож.стоимости");

            dt.Columns.Add("Код валюты контракта");
            dt.Columns.Add("Курс валюты");
            dt.Columns.Add("Дата курса валюты");
            dt.Columns.Add("Фактурная стоимость");
            dt.Columns.Add("Таможенная стоимость");
            dt.Columns.Add("Статистическая стоимость");
            dt.Columns.Add("Цена за кг., USD");
            dt.Columns.Add("Цена за ед. изм., USD");
            dt.Columns.Add("Цена за ед. изм., руб.");
            dt.Columns.Add("Соотн. нетто-брутто, %");

            foreach (MyRow myRow in myRows)
                dt.Rows.Add(myRow.Row);

            return dt;
        }
    }
}
