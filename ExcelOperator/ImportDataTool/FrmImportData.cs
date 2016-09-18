using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportDataTool
{
    public partial class FrmImportData : Form
    {
        public FrmImportData()
        {
            InitializeComponent();
        }
        string fileName = string.Empty;
        string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;data source=";
        DALLib dal = new DALLib("conn");

        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = @"F:\MyJianGuoYun\FOM";//注意这里写路径时要用c:\\而不是c:\
            openFileDialog.Filter = "Excel|*.xls|所有文件|*.*";
            openFileDialog.RestoreDirectory = true;
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
                        conn.Dispose();
                    }
                    if (excelSheets != null && excelSheets.Rows.Count > 0)
                    {
                        ArrayList rows = new ArrayList();
                        foreach (DataRow row in excelSheets.Rows)
                        {
                            string TableName = row["TABLE_NAME"].ToString();
                            //TableName.Contains("FilterDatabase")
                            if (TableName.Contains("$_") || TableName.Contains("$'_"))
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

        private void cboSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cboSheets.Text))
                return;
            string tableName = cboSheets.Text;
            // 查询语句
            string sql = "SELECT * FROM [" + tableName + "]";
            DataSet ds = new DataSet();
            OleDbDataAdapter da = new OleDbDataAdapter(sql, connStr);
            da.Fill(ds);

            if (ds.Tables.Count > 0)
            {
                dataGridView1.DataSource = ds.Tables[0];
                DataTable tab = ds.Tables[0];
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            string tableName = cboSheets.Text.Replace("$", "").Replace("(1)", "").Replace("(2)", "").Replace("(3)", "").Replace("(4)", "").Replace("'","");
            DataTable tab = dataGridView1.DataSource as DataTable;

            try
            {
                for (int i = 0; i < tab.Rows.Count; i++)
                {
                    string colName = string.Empty;
                    string colValue = string.Empty;
                    string InsertSql = string.Empty;
                    for (int j = 0; j < tab.Columns.Count; j++)
                    {
                        string rowValue = tab.Rows[i][j].ToString();
                        if (j == 0)
                        {
                            colName += tab.Columns[j].ColumnName;
                            colValue += rowValue.Contains("'") ? "'" + rowValue.Replace("'", "''") + "'" : "'" + rowValue + "'";
                        }
                        else
                        {
                            colName += "," + tab.Columns[j].ColumnName;
                            colValue += rowValue.Contains("'") ? ",'" + rowValue.Replace("'", "''") + "'" : ",'" + rowValue + "'";
                        }
                    }
                    InsertSql = string.Format("INSERT INTO {0} ({1})VALUES({2})", tableName, colName, colValue);

                    dal.Update(InsertSql);
                }
                MessageBox.Show("Import success！", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
