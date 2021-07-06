using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Data;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;

namespace Email_Report_With_Attachment
{
    public class EmailAttachments
    {
        static MemoryStream BindDataforPDF(int StateCode, string stateName, string fromDate)
        {

            // GET DATA FOR PARTICULAR DATE
            IMisc x;
            Hashtable param;
            x = new MiscCMD();
            param = new Hashtable();
            param.Clear();
            param.Add("@Key", "KeyValue");                        
            DataSet dset = x.GetData("SP_Name", "PORT", param, 0);

            StringWriter stringWrite = new StringWriter();
            DataTable dt = new DataTable();
            MemoryStream memoryStream = new MemoryStream();
            //Info line in PDF
            StringBuilder stringb = new StringBuilder();

            stringWrite = CreateHeaderFooter("Header", stringWrite);
            stringWrite = AddBlankLine(stringWrite, " ");

            stringWrite = AddBlankLine(stringWrite, "Report Card regarding Data Compliance and Project Implementation dated " + fromDate + " for " + stateName);
            stringWrite = AddBlankLine(stringWrite, " ");
            //1. Fetch DataTable
            stringWrite.WriteLine("1. Report");
            stringWrite = AddBlankLine(stringWrite, " ");
            //stringWrite = 
            if (dset.Tables[0].Rows.Count > 0)
            {
                dt = dset.Tables[0];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 1);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            // continue for next reports in repetition 


            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite = CreateHeaderFooter("Footer", stringWrite);

            //Finally Write the writer into a reader
            StringReader sr = new StringReader(stringWrite.ToString());
            //Generate the pdf
            Document pdfDoc = new Document(new Rectangle(288f, 144f), 10f, 10f, 10f, 0f);
            pdfDoc.SetPageSize(PageSize.A4.Rotate());
            HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
            var writer = PdfWriter.GetInstance(pdfDoc, memoryStream);
            pdfDoc.Open();
            htmlparser.Parse(sr);
            writer.CloseStream = false; // use when send email 
            pdfDoc.Close();
            return memoryStream; // Use when send email 

        }
        static StringWriter AddBlankLine(StringWriter stringWrite1, string Msz)
        {
            System.Web.UI.HtmlTextWriter htmlWrite1 = new HtmlTextWriter(stringWrite1);
            Panel panel = new Panel();
            Label label = new Label();
            label.Text = Msz;
            label.Font.Underline = true;
            label.Font.Size = 10;
            panel.Controls.Add(label);
            panel.RenderControl(htmlWrite1);            
            return stringWrite1;
        }
        protected static StringWriter CreateHeaderFooter(string type, StringWriter stringWriter)
        {
            System.Web.UI.HtmlTextWriter htmlWrite1 = new HtmlTextWriter(stringWriter);
            Panel panel = new Panel();
            if (type == "Header")
            {
                HtmlImage htmlImage = new HtmlImage();
                htmlImage.Src = "URL of Image";
                htmlImage.Width = 120;
                htmlImage.Height = 120;
                htmlImage.Align = "Center";
                panel.Controls.Add(htmlImage);
                Label label = new Label();
                label.Text = "Heading Value";
                label.Font.Underline = true;
                label.Font.Size = 15;
                label.Attributes.Add("style", "text-align:center");
                panel.Controls.Add(label);
                panel.RenderControl(htmlWrite1);
                //stringWrite1(htmlWrite1);
            }
            else if (type == "Footer")
            {
                Label label = new Label();
                label.Text = "Developed by ********";
                label.Font.Underline = true;
                label.Font.Size = 10;
                label.Attributes.Add("style", "text-align:center");
                panel.Controls.Add(label);
                panel.RenderControl(htmlWrite1);
            }
            return stringWriter;
        }
        protected static StringWriter getStringWriterbyDataTable(DataTable dt, StringWriter stringWrite1, int tableNo)
        {
            DataTable dataTable = new DataTable();
            //System.IO.StringWriter stringWrite1 = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter htmlWrite1 = new HtmlTextWriter(stringWrite1);
            DataGrid dataGrid = new DataGrid();
            dataGrid.DataSource = dt;
            dataGrid.DataBind();
            dataGrid.HeaderStyle.BackColor = System.Drawing.Color.DarkOrange;
            dataGrid.HeaderStyle.ForeColor = System.Drawing.Color.White;
            dataGrid.HeaderStyle.Font.Bold = true;
            dataGrid.RenderControl(htmlWrite1);
            return stringWrite1;
        }
        public static Attachment GetAttachment(int StateCode, string StateName, string fromDate)
        {
            var file = BindDataforPDF(StateCode, StateName, fromDate);
            file.Seek(0, SeekOrigin.Begin);
            // Attachment attachment = new Attachment(new MemoryStream(), "RNATeam.pdf", "application/pdf");
            string attachmentName = StateName + "_" + "PDFFileName.pdf";
            Attachment attachment = new Attachment(file, attachmentName, "application/pdf");
            ContentDisposition disposition = attachment.ContentDisposition;
            disposition.CreationDate = System.DateTime.Now;
            disposition.ModificationDate = System.DateTime.Now;
            disposition.DispositionType = DispositionTypeNames.Attachment;
            return attachment;
        }       
    }
}
