using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Analysis
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            loadData();

            MyTable myTable = MyTable.getInstance();

            dgvMain.DataSource = myTable.ToDataTable();
        }

        private void loadData()
        {
            prBar.Visible = true;
            prBar.Minimum = 2;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xls)|*.xls";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Multiselect = true;

            MyTable myTable = MyTable.getInstance();
            DataTable dt = new DataTable();
            for (int i = 0; i < 62; i++ )
                dt.Columns.Add(i.ToString());

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in openFileDialog.FileNames)
                {
                    ExcelDoc excelDoc = new ExcelDoc(fileName);

                    prBar.Maximum = excelDoc.getCount("A", 2);                    
                    prBar.Value = prBar.Minimum;

                    int i = 2;

                    while (excelDoc.getValue("A" + i.ToString(), "A" + i.ToString()) != null)
                    {
                        dt.Rows.Add(excelDoc.getRow(i));

                        i++;
                        prBar.Value++;
                    }

                    excelDoc.Dispose();
                }

                myTable.loadData(dt);
            }

            prBar.Visible = false;
            prBar.Value = prBar.Minimum;
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            string[] keys = tbKeys.Text.Split(',');

            MyTable myTable = MyTable.getInstance();

            dgvMain.DataSource = myTable.ToDataTable(keys);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyTable myTable = MyTable.getInstance();
            myTable.clearList();

            dgvMain.DataSource = null;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            prBar.Visible = true;
            prBar.Minimum = 0;
            prBar.Value = prBar.Minimum;
            prBar.Maximum = dgvMain.Rows.Count;

            ExcelDoc excelDoc = new ExcelDoc();

            excelDoc.setValue(1, 1, "Номер декларации");
            excelDoc.setValue(1, 2, "№ товара");
            excelDoc.setValue(1, 3, "Таможня");
            excelDoc.setValue(1, 4, "Дата");
            excelDoc.setValue(1, 5, "Направление");
            excelDoc.setValue(1, 6, "Режим");
            excelDoc.setValue(1, 7, "Наименование отправителя");
            excelDoc.setValue(1, 8, "ИНН отправителя");
            excelDoc.setValue(1, 9, "Адрес отправителя");
            excelDoc.setValue(1, 10, "Наименование получателя");
            excelDoc.setValue(1, 11, "ИНН получателя");
            excelDoc.setValue(1, 12, "Адрес получателя");
            excelDoc.setValue(1, 13, "Наименование декларанта");
            excelDoc.setValue(1, 14, "ИНН декларанта");
            excelDoc.setValue(1, 15, "Адрес декларанта");
            excelDoc.setValue(1, 16, "Наименований");
            excelDoc.setValue(1, 17, "Количество мест");
            excelDoc.setValue(1, 18, "Вид декларации");
            excelDoc.setValue(1, 19, "Код страны нахождения");
            excelDoc.setValue(1, 2, "Код торгующей страны");
            excelDoc.setValue(1, 21, "Код страны декл.");
            excelDoc.setValue(1, 22, "Код страны отпр.");
            excelDoc.setValue(1, 23, "Код страны назнач.");
            excelDoc.setValue(1, 24, "Условие поставки");
            excelDoc.setValue(1, 25, "Пункт поставки");
            excelDoc.setValue(1, 26, "Код таможни на границе");

            excelDoc.setValue(1, 27, "Наименование изготовителя");
            excelDoc.setValue(1, 28, "Товарный знак");
            excelDoc.setValue(1, 29, "Код ТН ВЭД");
            excelDoc.setValue(1, 30, "Описание и характеристика товара");
            excelDoc.setValue(1, 31, "Упаковка");
            excelDoc.setValue(1, 32, "Количество товара");
            excelDoc.setValue(1, 33, "Единица измерения");
            excelDoc.setValue(1, 34, "Количество товара");
            excelDoc.setValue(1, 35, "Единица измерения");
            excelDoc.setValue(1, 36, "Код единицы измерения");
            excelDoc.setValue(1, 37, "Количество товара");
            excelDoc.setValue(1, 38, "Единица измерения");
            excelDoc.setValue(1, 39, "Код единицы измерения");
            excelDoc.setValue(1, 40, "Код страны происх.");
            excelDoc.setValue(1, 41, "Преференции");
            excelDoc.setValue(1, 42, "КТС");
            excelDoc.setValue(1, 43, "Метод");
            excelDoc.setValue(1, 44, "Платёж (дол)");
            excelDoc.setValue(1, 45, "Платёж (руб)");
            excelDoc.setValue(1, 46, "Фамилия");
            excelDoc.setValue(1, 47, "Телефон");
            excelDoc.setValue(1, 48, "Тип КТС");
            excelDoc.setValue(1, 49, "Должность");
            excelDoc.setValue(1, 50, "Вес нетто (кг)");
            excelDoc.setValue(1, 51, "Вес брутто (кг)");
            excelDoc.setValue(1, 52, "Код валюты тамож.стоимости");

            excelDoc.setValue(1, 53, "Код валюты контракта");
            excelDoc.setValue(1, 54, "Курс валюты");
            excelDoc.setValue(1, 55, "Дата курса валюты");
            excelDoc.setValue(1, 56, "Фактурная стоимость");
            excelDoc.setValue(1, 57, "Таможенная стоимость");
            excelDoc.setValue(1, 58, "Статистическая стоимость");
            excelDoc.setValue(1, 59, "Цена за кг., USD");
            excelDoc.setValue(1, 60, "Цена за ед. изм., USD");
            excelDoc.setValue(1, 61, "Цена за ед. изм., руб.");
            excelDoc.setValue(1, 62, "Соотн. нетто-брутто, %");

            int i = 2;

            foreach (DataGridViewRow row in dgvMain.Rows)
            {
                for (int j = 0; j < row.Cells.Count; j++)
                    excelDoc.setValue(i, j + 1, row.Cells[j].Value.ToString());

                prBar.Value++;

                i++;
            }

            prBar.Visible = false;

            excelDoc.Show();
        }
    }
}
