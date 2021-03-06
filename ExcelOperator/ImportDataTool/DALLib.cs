﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataTool
{
    public class DALLib
    {
        string connstr = string.Empty;
        SqlConnection conn = new SqlConnection();
        SqlCommand cmd = new SqlCommand();

        public DALLib(string name)
        {
            connstr = ConfigurationManager.ConnectionStrings[name].ToString();
            conn = new SqlConnection(connstr);
        }

        public bool Update(string sqlstr)
        {
            bool result = false;
            try
            {
                conn.Open();
                cmd = new SqlCommand(sqlstr, conn);
                result = cmd.ExecuteNonQuery() > 0 ? true : false;
                conn.Close();
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                conn.Close();
            }
            return result;
        }
    }
}
