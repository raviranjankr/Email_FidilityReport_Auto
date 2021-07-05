using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Net.Mail;
using System.Collections;

namespace Email_FidilityReport_Auto
{
    class Program
    {
        private static string SendEmail(Attachment attachment, string emailContent, string emailDate, string emailID, string emailCC)
        {
            string recipient = emailID;
            string ccList = emailCC;
            return new EmailProvider().Send("PMAYMIS"
                    , "pmaymis-mhupa@gov.in"
                    , recipient
                    , recipient
                    , recipient
                    , "PMAYMIS : Report Card regarding Data Compliance and Project Implementation dated " + emailDate
                    , emailContent
                    , EmailProvider.Format.Html
                    , attachment, ccList, null);
        }
        static void Main(string[] args)
        {
            StringBuilder strLOG = new StringBuilder();

            if (DateTime.Now.DayOfWeek.ToString().ToLower() == "monday")
            {
                LOG.WriteToLog("******** Service Started for sending auto email of fidility report ********");
                try
                {
                    IMisc x;
                    Hashtable param;
                    x = new MiscCMD();
                    param = new Hashtable();
                    param.Clear();
                    param.Add("@Status", 0);
                    DataSet ds = x.GetData("FidilityReport_For_Email", "22", param, 0);
                    if (ds.Tables.Count == 2)
                    {
                        int stateCode = 0;
                        string state_Name = "";
                        string emailID_To = "";
                        string emailID_CC = "";
                        string DOE = "";

                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                LOG.WriteToLog("EMAIL IDs, CCs and State Records : CHECKED");
                              
                                stateCode = Convert.ToInt32(ds.Tables[0].Rows[i]["State_Code"].ToString());                        
                                state_Name = ds.Tables[0].Rows[i]["State_Name"].ToString();
                                emailID_To = ds.Tables[0].Rows[i]["EmailID_TO"].ToString();
                                emailID_CC = ds.Tables[0].Rows[i]["EmailID_CC"].ToString();
                                DOE = ds.Tables[1].Rows[0]["DOE"].ToString();

                                EmailTemplate emailTemplate = new EmailTemplate();
                                emailTemplate.REPORT_DATE = DOE;

                                var emailContent = emailTemplate.GetEmailContent();
                                //Step 2: Generate the PDF and hold in memory
                                var attachment = EmailAttachments.GetAttachment(stateCode, state_Name, DOE);
                                //Step 3: Send Email
                                var sendEmail = SendEmail(attachment, emailContent, DOE, emailID_To, emailID_CC);
                                if (sendEmail.ToString() == "Email sent successfully")
                                {
                                    LOG.WriteToLog("STATUS : EMAIL SENT FOR STATE CODE : " + stateCode);
                                    
                                }
                                else
                                {
                                    LOG.WriteToLog("STATUS : EMAIL NOT SENT FOR STATE CODE : " + stateCode);                                    
                                }                                
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LOG.WriteToLog("EXCEPTION :  " + ex.Message.ToString());                    
                }
            }
        }
    }
}
