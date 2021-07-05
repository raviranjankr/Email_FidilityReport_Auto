using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.Configuration;

namespace Email_FidilityReport_Auto
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
        public int InsertUpdate(string procName, Hashtable hashParam)
        {
            try
            {
                int res = 0;
                OpenConnection("22");
                mCom.CommandText = procName;
                mCom.CommandType = CommandType.StoredProcedure;
                foreach (DictionaryEntry de in hashParam)
                    mCom.Parameters.AddWithValue(de.Key.ToString(), de.Value);
                res = mCom.ExecuteNonQuery();
                mCom.Parameters.Clear();
                CloseConnection();
                DisposeConnection();
                return res;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }
        public int InsertUpdateWith_SQLOut(string procName, Hashtable hashParam, Hashtable outParam, out string[] respSQLOut)
        {
            int isExecuted = 0;
            string sqlOUTValue = "";
            respSQLOut = new string[2];
            try
            {
                miscModel = new MiscModel();
                OpenConnection("22");
                mCom.CommandText = procName;
                mCom.CommandType = CommandType.StoredProcedure;
                foreach (DictionaryEntry de in hashParam)
                    mCom.Parameters.AddWithValue(de.Key.ToString(), de.Value);

                string[] outPramaterColumnNames = new string[2];
                foreach (DictionaryEntry de in outParam)
                {
                    if (de.Value.ToString() == "Int")
                    {
                        outPramaterColumnNames[0] = de.Key.ToString();
                        mCom.Parameters.Add(de.Key.ToString(), SqlDbType.Int, 10);
                        mCom.Parameters[de.Key.ToString()].Direction = ParameterDirection.Output;
                    }
                    if (de.Value.ToString() == "NVarChar")
                    {
                        outPramaterColumnNames[1] = de.Key.ToString();
                        mCom.Parameters.Add(de.Key.ToString(), SqlDbType.NVarChar, 50).Direction = ParameterDirection.Output;
                        mCom.Parameters[de.Key.ToString()].Direction = ParameterDirection.Output;
                    }
                }
                mCom.ExecuteNonQuery();
                if (outParam.Count == 1)
                {

                    isExecuted = Convert.ToInt16(mCom.Parameters[outPramaterColumnNames[0].ToString()].Value.ToString());
                    sqlOUTValue = "";
                    respSQLOut[0] = isExecuted.ToString();
                    respSQLOut[1] = sqlOUTValue;
                }
                else if (outParam.Count == 2)
                {
                    string str = mCom.Parameters["@AppID"].Value.ToString();
                    isExecuted = Convert.ToInt16(mCom.Parameters["@AppID"].Value.ToString());
                    sqlOUTValue = mCom.Parameters[outPramaterColumnNames[1]].Value.ToString();
                    respSQLOut[0] = isExecuted.ToString();
                    respSQLOut[1] = sqlOUTValue;
                }
                mCom.Parameters.Clear();
                CloseConnection();
                DisposeConnection();
                return isExecuted;
            }
            catch (Exception ex)
            {
                return isExecuted;
            }
        }
    }
}
