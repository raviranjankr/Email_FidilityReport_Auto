using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Data;

namespace Email_FidilityReport_Auto
{
    public interface IMisc
    {
        int InsertUpdate(string procName, Hashtable hashParam);
        int InsertUpdateWith_SQLOut(string procName, Hashtable hashParam, Hashtable outParam, out string[] sqlOUT);
        DataSet GetData(string procName, string conType, Hashtable hashParam, int timeOut);
        DataSet GetDataWith_SQLOut(string procName, string conType, Hashtable hashParam, int timeOut, out int TotalRecCount);
    }
}