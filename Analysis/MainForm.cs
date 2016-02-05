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
            openFileDialog.Filter = "Excel files (*.xls, *.xlsx)|*.xls; *.xlsx";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Multiselect = true;

            MyTable myTable = MyTable.getInstance();
            

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in openFileDialog.FileNames)
                {
                    using (ExcelDoc excelDoc = new ExcelDoc(fileName))
                    {
                        string columnName = "B";
                        string columnName2 = "C";

                        prBar.Maximum = excelDoc.getCount(columnName, columnName2, 2);
                        prBar.Value = prBar.Minimum;

                        int i = 2;

                        while ((excelDoc.getValue(columnName + i.ToString(), columnName + i.ToString()) != null) || (excelDoc.getValue(columnName + (i + 1).ToString(), columnName + (i + 1).ToString()) != null)
                            || (excelDoc.getValue(columnName2 + i.ToString(), columnName2 + i.ToString()) != null) || (excelDoc.getValue(columnName2 + (i + 1).ToString(), columnName2 + (i + 1).ToString()) != null))
                        {
                            if ((excelDoc.getValue(columnName + i.ToString(), columnName + i.ToString()) == null) && (excelDoc.getValue(columnName2 + i.ToString(), columnName2 + i.ToString()) == null))
                                i++;

                            myTable.Add(new MyRow(excelDoc.getRow(i)));

                            i++;
                            prBar.Value++;
                        }
                    }
                }
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

            excelDoc.setValue(1, 1, "Артикул");
            excelDoc.setValue(1, 2, "Параметр");
            excelDoc.setValue(1, 3, "Описание");
            
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
