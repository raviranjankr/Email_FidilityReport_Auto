using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace Email_FidilityReport_Auto
{
    public class MiscModel
    {
        public int OutPutParamterValue1 { get; set; }
        public string OutPutParamterValue2 { get; set; }
        public int TotalRecCount { get; set; }
        public int Status_Int { get; set; }

        //
        public int State_Code { get; set; }
        public int District_Code { get; set; }
        public int City_Code { get; set; }       
    }
}