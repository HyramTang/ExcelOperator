using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelOperator
{
    public partial class FrmExcelReader : Form
    {
        public FrmExcelReader()
        {
            InitializeComponent();
        }
        string fileName = string.Empty;
        string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;data source=";

        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";//注意这里写路径时要用c:\\而不是c:\
            openFileDialog.Filter = "Excel|*.xls|所有文件|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog.FileName;
                txtAddress.Text = fileName;

                if (!string.IsNullOrEmpty(fileName))
                {
                    //连接字符串
                    connStr += fileName;
                    DataTable excelSheets = new DataTable();
                    using (OleDbConnection conn = new OleDbConnection(connStr))
                    {
                        conn.Open();
                        excelSheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                        conn.Close();
                    }
                    if (excelSheets != null && excelSheets.Rows.Count > 0)
                    {
                        ArrayList rows = new ArrayList();
                        foreach (DataRow row in excelSheets.Rows)
                        {
                            string TableName = row["TABLE_NAME"].ToString();
                            if (TableName.Contains("FilterDatabase"))
                                rows.Add(row);
                        }
                        foreach (DataRow row in rows)
                        {
                            excelSheets.Rows.Remove(row);
                        }
                        cboSheets.DisplayMember = "TABLE_NAME";
                        cboSheets.DataSource = excelSheets;
                    }
                }
            }
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            string tableName = cboSheets.Text;
            rtxtOutput.Text = string.Empty;
            // 查询语句
            string sql = "SELECT * FROM [" + tableName + "]";
            DataSet ds = new DataSet();
            OleDbDataAdapter da = new OleDbDataAdapter(sql, connStr);
            da.Fill(ds);

            if (ds.Tables.Count > 0)
            {
                dataGridView1.DataSource = ds.Tables[0];
                DataTable tab = ds.Tables[0];
                for (int i = 0; i < tab.Rows.Count; i++)
                {
                    string strA = tab.Rows[i][0].ToString();
                    string strB = tab.Rows[i][1].ToString();
                    string[] arrayB = strB.Split(';');
                    if (arrayB.Length > 0)
                    {
                        foreach (string str in arrayB)
                        {
                            rtxtOutput.AppendText(strA + "\t" + str + "\r");
                        }
                    }
                    else
                        rtxtOutput.AppendText(strA + "\t" + strB + "\r");
                }
            }
        }
    }
}
