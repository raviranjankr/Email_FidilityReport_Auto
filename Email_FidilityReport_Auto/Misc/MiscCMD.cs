using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.Configuration;

namespace Email_Report_With_Attachment
{
    public class MiscCMD : IMisc
    {
        //Hashtable hashParam;
        SqlConnection conn;
        SqlCommand mCom;
        SqlDataAdapter mDa;
        MiscModel miscModel;
        #region OpenCloseDispose_Connection
        private void OpenConnection(string conType)
        {
            string con = "";
            if (conType == "23")
            { con = "con"; }
            else if (conType == "21")
            { con = "con"; }
            else
            { con = "con"; }

            conn = new SqlConnection(ConfigurationManager.AppSettings[con].ToString());
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            mCom = new SqlCommand();
            mCom.Connection = conn;
        }
        private void CloseConnection()
        {
            if (conn == null == false)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
        private void DisposeConnection()
        {
            if (conn == null == false)
                conn.Dispose();
        }
        #endregion
        public DataSet GetData(string procName, string conType, Hashtable hashParam, int timeOut)
        {
            DataSet ds = new DataSet();
            try
            {
                OpenConnection(conType);
                mCom.CommandText = procName;
                mCom.CommandType = CommandType.StoredProcedure;

                mDa = new SqlDataAdapter(mCom);
                foreach (DictionaryEntry de in hashParam)
                    mCom.Parameters.AddWithValue(de.Key.ToString(), de.Value);

                mDa.SelectCommand.CommandTimeout = timeOut;
                mDa.Fill(ds);
                mCom.Parameters.Clear();
                CloseConnection();
                DisposeConnection();
            }
            catch (Exception ex) { }
            return ds;
        }
        public DataSet GetDataWith_SQLOut(string procName, string conType, Hashtable hashParam, int timeOut, out int TotalRecCount)
        {
            DataSet ds = new DataSet();
            TotalRecCount = 0;
            try
            {
                OpenConnection(conType);
                mCom.CommandText = procName;
                mCom.CommandType = CommandType.StoredProcedure;

                foreach (DictionaryEntry de in hashParam)
                    mCom.Parameters.AddWithValue(de.Key.ToString(), de.Value);
                mCom.Parameters.Add("@RecordCount", SqlDbType.Int, 4);
                mCom.Parameters["@RecordCount"].Direction = ParameterDirection.Output;
                mCom.CommandTimeout = timeOut;

                using (SqlDataAdapter da = new SqlDataAdapter(mCom))
                {
                    da.Fill(ds);
                }
                if (mCom.Parameters["@RecordCount"].Value.ToString() != "")
                {
                    TotalRecCount = Convert.ToInt32(mCom.Parameters["@RecordCount"].Value);
                }
                else
                {
                    TotalRecCount = Convert.ToInt32(!Convert.IsDBNull(mCom.Parameters["@RecordCount"].Value));
                }
                mCom.Parameters.Clear();
                CloseConnection();
                DisposeConnection();
            }
            catch (Exception ex)
            {
                ds = new DataSet();
            }
            return ds;
        }
        
    }
}
