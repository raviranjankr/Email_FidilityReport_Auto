using System.Text;

namespace Email_Report_With_Attachment
{
    public class EmailTemplate
    {
        StringBuilder sbEmailMsg;
        public string REPORT_DATE = "";
        public string GetEmailContent()
        {
            sbEmailMsg = new StringBuilder();
            ComposeEmail();
            return sbEmailMsg.ToString();
        }

        #region Private methods

        private void ComposeEmail()
        {
            EmailHeader();
            EmailBody();
            EmailFooter();
        }

        private void EmailHeader()
        {
            sbEmailMsg.Append("<html><head><body><left>Dear Sir/Madam,</left>");
        }

        private void EmailBody()
        {
            //REPORT_DATE
            sbEmailMsg.Append("<p>Kindly find enclosed herewith Report Card regarding Data Compliance and Project Implementation dated " + REPORT_DATE + " which reflects the overall pendency in MIS and Geo-tagging with the progress made so far for your urgent necessary action. </p>");
        }

        private void EmailFooter()
        {
            sbEmailMsg.Append("<br/><br/><br/><br/><b>Thanks & Regards</b>,<br/> Organization Name,<br/> Organizational Body</body></html>");
        }
        #endregion
    }
}
