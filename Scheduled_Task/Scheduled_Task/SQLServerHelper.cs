using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scheduled_Task
{
    class SQLServerHelper
    {
        public static string ConnString = string.Empty;

        public static string Conn_Config_Str_Name = string.Empty;

        public static string Conn_Server = ConfigurationManager.AppSettings["Conn_Server"];
        public static string Conn_DBName = ConfigurationManager.AppSettings["Conn_DBName"];
        public static string Conn_Uid = ConfigurationManager.AppSettings["Conn_Uid"];
        public static string Conn_Pwd = ConfigurationManager.AppSettings["Conn_Pwd"];

        private static string _ConnString
        {
            get
            {
                if (!string.IsNullOrEmpty(ConnString))
                    return (ConnString);

                object oConn = ConfigurationManager.ConnectionStrings[Conn_Config_Str_Name];
                if (oConn != null && oConn.ToString() != "")
                    return (oConn.ToString());

                return (string.Format(@"server={0};database={1};uid={2};password={3}", Conn_Server, Conn_DBName, Conn_Uid, Conn_Pwd));
            }
        }

        /* 取datatable */
        public static DataTable GetDataTable(out string sError, string sSQL)
        {
            DataTable dt = null;
            sError = string.Empty;

            try
            {
                SqlConnection conn = new SqlConnection(_ConnString);
                SqlCommand comm = new SqlCommand();
                comm.Connection = conn;
                comm.CommandText = sSQL;
                SqlDataAdapter dapter = new SqlDataAdapter(comm);
                dt = new DataTable();
                dapter.Fill(dt);
            }
            catch (Exception ex)
            {
                sError = ex.Message;
            }

            return (dt);
        }


        /* 取dataset */
        public static DataSet GetDataSet(out string sError, string sSQL)
        {
            DataSet ds = null;
            sError = string.Empty;

            try
            {
                SqlConnection conn = new SqlConnection(_ConnString);
                SqlCommand comm = new SqlCommand();
                comm.Connection = conn;
                comm.CommandText = sSQL;
                SqlDataAdapter dapter = new SqlDataAdapter(comm);
                ds = new DataSet();
                dapter.Fill(ds);
            }
            catch (Exception ex)
            {
                sError = ex.Message;
            }

            return (ds);
        }


        /* 取某个单一的元素 */
        public static object GetSingle(out string sError, string sSQL)
        {
            DataTable dt = GetDataTable(out sError, sSQL);
            if (dt != null && dt.Rows.Count > 0)
            {
                return (dt.Rows[0][0]);
            }

            return (null);
        }


        /* 取最大的ID */
        public static Int32 GetMaxID(out string sError, string sKeyField, string sTableName)
        {
            DataTable dt = GetDataTable(out sError, "select isnull(max([" + sKeyField + "]),0) as MaxID from [" + sTableName + "]");
            if (dt != null && dt.Rows.Count > 0)
            {
                return (Convert.ToInt32(dt.Rows[0][0].ToString()));
            }

            return (0);
        }


        /* 执行 insert,update,delete 动作，也可以使用事务 */
        public static bool UpdateData(out string sError, string sSQL, bool bUseTransaction = false)
        {
            int iResult = 0;
            sError = string.Empty;

            if (!bUseTransaction)
            {
                try
                {
                    SqlConnection conn = new SqlConnection(_ConnString);
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    SqlCommand comm = new SqlCommand();
                    comm.Connection = conn;
                    comm.CommandText = sSQL;
                    iResult = comm.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    sError = ex.Message;
                    iResult = -1;
                }
            }
            else
            { /* 使用事务 */
                SqlTransaction trans = null;
                try
                {
                    SqlConnection conn = new SqlConnection(_ConnString);
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    trans = conn.BeginTransaction();
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandText = sSQL;
                    cmd.Transaction = trans;
                    iResult = cmd.ExecuteNonQuery();
                    trans.Commit();
                }
                catch (Exception ex)
                {
                    sError = ex.Message;
                    iResult = -1;
                    trans.Rollback();
                }
            }

            return (iResult > 0);
        }
    }
}
