using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Specialized;

namespace WindowsFormsApp1
{
    
    public partial class Form1 : Form
    {
        Excel.Application excelApp = new Excel.ApplicationClass();
        NameValueCollection vdata = new NameValueCollection();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string url = "C:/test/playauto_emp_site_code_23.xls";
            try
            {
                Excel.Workbook wb = excelApp.Workbooks.Open(url);
                Excel.Sheets _Sheets = wb.Sheets;
                //첫번째 worksheet 선택
                Excel.Worksheet ws = _Sheets[1] as Excel.Worksheet;

                object chkcell1 = ws.Range["A1"].Value;

                if (chkcell1.ToString().Equals("쇼핑몰코드표"))
                {
                    //열들에 data를 배열로 받아옴
                    foreach (Excel.Range _row in ws.UsedRange.Rows)
                    {
                        Object[,] edata = (System.Object[,])_row.Value;
                        if (edata[1, 2] != null)
                        {
                            vdata.Add((string)edata[1, 1], (string)edata[1, 2]);
                        }
                    }
                    SetupGridView(vdata);
                }
                else
                {
                    throw new Exception("양식 오류입니다.");
                }

                wb.Close();
                ws = null;
                _Sheets = null;
            }
            catch (Exception ee)
            {
                throw new Exception(ee.Message); 
            }

        }

        private void SetupGridView(NameValueCollection vdata)
        {
            foreach (string fieldName in vdata.Keys)
            {
                foreach (string fieldValue in vdata.GetValues(fieldName))
                {
                    if (!(fieldValue == ""))
                    { 
                        string[] row = { fieldName, fieldValue };
                        dataGridView1.Rows.Add(row);
                    }
                }
            }
        }

        private void SetupGridView(String searchKey)
        {
                var searchList = from key in vdata.Keys.Cast<String>()
                                 from value in vdata.GetValues(key)
                                 where key.Contains(searchKey)
                                 select new { key, value };

                foreach (var test in searchList)
                {
                    string[] row = { test.key, test.value };
                    dataGridView2.Rows.Add(row);
                }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void searchSite(object sender, EventArgs e)
        {
              if (dataGridView2.Rows.Count > 1)
            {
                RemoveAllRows();
            }
            SetupGridView(searchText.Text);
        }

        private void RemoveAllRows()
        {
            for(int i=0; i<dataGridView2.Rows.Count-1;)
            {
                dataGridView2.Rows.Remove(dataGridView2.Rows[i]);
            }
        }

        private void form_load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].Name = "쇼핑몰";
            dataGridView1.Columns[1].Name = "코드";

            dataGridView2.ColumnCount = 2;
            dataGridView2.Columns[0].Name = "쇼핑몰";
            dataGridView2.Columns[1].Name = "코드";
        }
    }
}
