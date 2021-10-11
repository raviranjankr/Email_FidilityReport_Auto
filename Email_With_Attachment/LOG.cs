using System;
using System.IO;

namespace Email_Report_With_Attachment
{
    class LOG
    {
        public static void WriteToLog(string text)
        {
            string filepath = @"C:\LOGS\Email_" + DateTime.Now.ToShortDateString().Replace("/", "").Replace("-", "") + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(text);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(text);
                }
            }
        }
    }
}
