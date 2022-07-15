using Focus.DatabaseFactory;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;

namespace SOAReport.Models
{
    public class DBClass
    {
        static SqlConnection con;
        public DBClass(string FSerName,string FSQLUID,string FSQLPWD)
        {
            string FDB = "Focus8Erp";
            string Fconnection = "";
            Fconnection = $"Server={FSerName};Database={FDB};User Id={FSQLUID};Password={FSQLPWD};";
            con = new SqlConnection(Fconnection);
        }
        public static int GetExecute(string strInsertOrUpdateQry, int CompId, ref string error)
        {
            try
            {
                Database obj = Focus.DatabaseFactory.DatabaseWrapper.GetDatabase(CompId);
                return (obj.ExecuteNonQuery(CommandType.Text, strInsertOrUpdateQry));
            }
            catch (Exception e)
            {
                SetLog(DateTime.Now.ToString() + " GetDataSet :" + e.Message);
                return 0;
            }
        }
        public static DataSet GetData(string strSelQry, int CompId, ref string error)
        {
            try
            {
                Database obj = Focus.DatabaseFactory.DatabaseWrapper.GetDatabase(CompId);
                return (obj.ExecuteDataSet(CommandType.Text, strSelQry));
            }
            catch (Exception e)
            {
                SetLog(DateTime.Now.ToString() + " GetDataSet :" + e.Message);
                return null;
            }
        }
        public DataSet GetData(string Query)
        {
            con.Open();
            SqlCommand cmd = new SqlCommand(Query, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            DataSet dst = ds;
            con.Close();
            return dst;
        }
        public static DataSet GetDataSet(string strselQry, int companyId, ref string logText)
        {
            DataSet dataset = null;
            try
            {
                Database _db = DatabaseWrapper.GetDatabase(companyId);

                using (var con = _db.CreateConnection())
                {
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.CommandText = strselQry;
                        cmd.CommandTimeout = 0;
                        dataset = _db.ExecuteDataSet(cmd);
                    }
                }


                return dataset;

            }
            catch (Exception e)
            {
                SetLog(DateTime.Now.ToString() + " GetDataSet :" + e.Message);
                return null;
            }

        }


        public static void SetLog(string content)
        {
            StreamWriter objSw = null;
            try
            {
                string sFilePath = System.IO.Path.GetTempPath() + "SOAReport_EventLogs" + DateTime.Now.Date.ToString("ddMMyyyy") + ".txt";
                objSw = new StreamWriter(sFilePath, true);
                objSw.WriteLine(DateTime.Now.ToString() + " " + content + Environment.NewLine);

                string AppLocation = "";
                AppLocation = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData);
                string folderName = AppLocation + "\\LogFiles";
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName);
                }
                string sFilePath2 = folderName + "\\SOAReport_EventLogs-" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
                objSw = new StreamWriter(sFilePath2, true);
                objSw.WriteLine(DateTime.Now.ToString() + " " + content + Environment.NewLine);
            }
            catch (Exception ex)
            {
                //SetLog("Error -" + ex.Message);
            }
            finally
            {
                if (objSw != null)
                {
                    objSw.Flush();
                    objSw.Dispose();
                }
            }
        }
        public void ODBCConnection()
        {
            string ConnectionString = @" DSN =NWDSN";
            OdbcConnection conn = new OdbcConnection(ConnectionString);

            // use the SQL Query to get the customers data  
            OdbcCommand cmd = new OdbcCommand("Select * From customers", conn);
            conn.Open();
        }
    }
}