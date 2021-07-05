using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;

namespace Email_FidilityReport_Auto
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
            param.Add("@Status", 1);
            param.Add("@State_Code", StateCode);
            param.Add("@Date", DateTime.ParseExact(fromDate.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture));
            DataSet dset = x.GetData("FidilityReport_For_Email", "22", param, 0);

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
            stringWrite.WriteLine("1. Project Annexure");
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


            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("2. Category Gender Modification Required");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[1].Rows.Count > 0)
            {
                dt = dset.Tables[1];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 2);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }

            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("3. Beneficiary Data Entered by ULB");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[2].Rows.Count > 0)
            {
                dt = dset.Tables[2];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 3);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }

            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("4. Beneficiary Attachment with Aadhaar");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[3].Rows.Count > 0)
            {
                dt = dset.Tables[3];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 4);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }

            stringWrite = AddBlankLine(stringWrite, " ");
            //5. Fetch DataTable 
            stringWrite.WriteLine("5. Beneficiary Attachment without Aadhaar");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[4].Rows.Count > 0)
            {
                dt = dset.Tables[4];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 5);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }

            //6. Fetch DataTable 
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("6. Aadhaar Mismatch");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[5].Rows.Count > 0)
            {
                dt = dset.Tables[5];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 6);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }


            //7. Fetch DataTable 
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("7. Aadhaar Replication");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[6].Rows.Count > 0)
            {
                dt = dset.Tables[6];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 7);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }


            //8. Fetch DataTable 
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("8. Female Beneficiaries / Joint Ownership");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[7].Rows.Count > 0)
            {
                dt = dset.Tables[7];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 8);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("9. Beneficiaries with Mobile No");
            stringWrite = AddBlankLine(stringWrite, "");
            if (dset.Tables[8].Rows.Count > 0)
            {
                dt = dset.Tables[8];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 9);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("10. Correction of Duplicate Mobile Numbers");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[9].Rows.Count > 0)
            {
                dt = dset.Tables[9];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 10);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("11. Uploading of revised Annexure (approved by CSMC)");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[10].Rows.Count > 0)
            {
                dt = dset.Tables[10];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 11);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("12. Physical Progress of Projects");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[11].Rows.Count > 0)
            {
                dt = dset.Tables[11];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 12);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("13. Submission of Physical MPR (Projects)");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[12].Rows.Count > 0)
            {
                dt = dset.Tables[12];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 13);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("14. Status of Central Assistance (Rs. in Cr.)");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[13].Rows.Count > 0)
            {
                dt = dset.Tables[13];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 14);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("15. Status of State or UT Share");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[14].Rows.Count > 0)
            {
                dt = dset.Tables[14];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 15);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("16. Financial MPR (Projects)");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[15].Rows.Count > 0)
            {
                dt = dset.Tables[15];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 16);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("17. DBT Expenditure Reported by State/ UTs in DBT Entry against Central Assistance Released (Rs. in Cr.) ");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[16].Rows.Count > 0)
            {
                dt = dset.Tables[16];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 17);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("18. DBT Expenditure Reported by ULB in MPR against Central Assistance Released (Rs. in Cr.) ");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[17].Rows.Count > 0)
            {
                dt = dset.Tables[17];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 18);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("19. Status of UC Received (Rs. in Cr.)");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[18].Rows.Count > 0)
            {
                dt = dset.Tables[18];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 19);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("20. Demand Survey Validation");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[19].Rows.Count > 0)
            {
                dt = dset.Tables[19];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 20);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
            stringWrite = AddBlankLine(stringWrite, " ");
            stringWrite.WriteLine("21. Total Projects and Houses Geo-tagged");
            stringWrite = AddBlankLine(stringWrite, " ");
            if (dset.Tables[20].Rows.Count > 0)
            {
                dt = dset.Tables[20];
                stringWrite = getStringWriterbyDataTable(dt, stringWrite, 21);
            }
            else
            {
                stringWrite = AddBlankLine(stringWrite, "-- No RECORD FOUND --");
            }
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
                htmlImage.Src = "http://pmaymis.gov.in/images/Pmay_Logo.jpg";
                htmlImage.Width = 120;
                htmlImage.Height = 120;
                htmlImage.Align = "Center";
                panel.Controls.Add(htmlImage);
                Label label = new Label();
                label.Text = "Pradhan Mantri Awas Yojana-Housing for All (Urban)";
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
                label.Text = "Developed by NIC-MoHUA Division";
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
            string attachmentName = StateName + "_" + "DataComplianceReport.pdf";
            Attachment attachment = new Attachment(file, attachmentName, "application/pdf");
            ContentDisposition disposition = attachment.ContentDisposition;
            disposition.CreationDate = System.DateTime.Now;
            disposition.ModificationDate = System.DateTime.Now;
            disposition.DispositionType = DispositionTypeNames.Attachment;
            return attachment;
        }
       
    }
}
