using System.Collections;
using System.Data;

namespace Email_Report_With_Attachment
{
    public interface IMisc
    {        
        DataSet GetData(string procName, string conType, Hashtable hashParam, int timeOut);
        DataSet GetDataWith_SQLOut(string procName, string conType, Hashtable hashParam, int timeOut, out int TotalRecCount);
    }
}