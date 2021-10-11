using System;
using System.Net.Mail;

namespace Email_Report_With_Attachment
{
    public class EmailProvider
    {
        public enum Format { Text, Html };
        public static readonly string SMTP_SERVER = "**.**.**.**";
        public static readonly string SMTP_USER = "*******@gov.in";
        public static readonly string SMTP_PWD = "";


        MailMessage message;
        SmtpClient smtp;
        public string Send
        (
            string senderName,
            string senderAddress,
            string recipientName,
            string recipientAddress,
            string replyToAddress,
            string subject,
            string body,
            Format format,
            Attachment attachments = null,
            string ccList = null,
            string bccList = null
        )
        {
            string response = string.Empty; 
            var done = false; 
            int retryCount = 0; 
            GetMessageContent(
                                senderName,
                                senderAddress,
                                recipientAddress,
                                replyToAddress,
                                subject,
                                body,
                                format,
                                attachments,
                                ccList,
                                bccList,
                                out message,
                                out smtp
                             );
            return SendEmail(
                                out response,
                                ref done,
                                ref retryCount,
                                message,
                                smtp
                             );
        }

        #region private methods
        
        private static void GetMessageContent(string senderName, string senderAddress, string recipientAddress, string replyToAddress, string subject, string body, Format format, Attachment attachments, string ccList, string bccList, out MailMessage message, out SmtpClient smtp)
        {
            string[] toEmail = recipientAddress.Split(',');
            message = new MailMessage();

            GetFromAddress(senderName, senderAddress, message);
            GetToAddressList(message, toEmail);
            GetCCAddressList(ccList, message);
            GetBCCAddressList(bccList, message);
            GetReplyToAddressList(replyToAddress, message);
            GetAttachmentList(attachments, message);

            message.Subject = subject;
            message.Body = body;
            message.IsBodyHtml = (format == Format.Html);

            // Port = Convert.ToInt32("25"), 25 for PMAY, 587 for gmail
            smtp = new SmtpClient
            {
                Host = SMTP_SERVER,
                Port = Convert.ToInt32("25"),
                EnableSsl = false, // false when using mission project
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false

            };
            // smtp.Credentials = new NetworkCredential(SMTP_USER, SMTP_PWD); // comment for mission project
        }
   

        private static void GetFromAddress(string senderName, string senderAddress, MailMessage message)
        {
            message.From = new MailAddress(senderAddress, senderName);
        }
 
        private static void GetToAddressList(MailMessage message, string[] toEmail)
        {
            foreach (string address in toEmail)
            {
                message.To.Add(new MailAddress(address.Trim()));
            }
        }
   
        private static void GetCCAddressList(string ccList, MailMessage message)
        {
            string[] ccEamil = null;
            if (ccList != null)
            {
                ccEamil = ccList.Split(',');
                foreach (string s in ccEamil)
                {
                    message.CC.Add(new MailAddress(s));
                }
            }
        }
    
        private static void GetBCCAddressList(string bccList, MailMessage message)
        {
            string[] bccEmail = null;
            if (bccList != null)
            {
                bccEmail = bccList.Split(',');
                foreach (string s in bccEmail)
                {
                    message.Bcc.Add(new MailAddress(s));
                }
            }
        }
        
        private static void GetReplyToAddressList(string replyToAddress, MailMessage message)
        {
            if (replyToAddress != null && replyToAddress.Length > 0)
            {
                message.Headers.Add("Reply-To", replyToAddress);
            }
        }        
        private static void GetAttachmentList(Attachment attachments, MailMessage message)
        {
            if (attachments != null)
            {
                message.Attachments.Add(attachments);
            }
        }
        
        private static string SendEmail(out string response, ref bool done, ref int retryCount, MailMessage message, SmtpClient smtp)
        {
            do
            {
                try
                {
                    smtp.Send(message);
                    done = true;
                    response = "Email sent successfully";
                }
                catch (SmtpException smtpEx)// when (retryCount++ <= 5)
                {
                    response = smtpEx.Message;
                }
                catch (Exception ex)// when (retryCount++ <= 5)
                {
                    response = ex.Message;
                }
            } while (!done);

            return response;
        }

        #endregion
    }
}
