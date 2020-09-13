using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace FinalDBToCSV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Vijay made all the changes , Vinisha just created the project. She did not do much on this.
            StringBuilder sb = new StringBuilder();
            DataTable dt = GetData();
            //this is for columns display
            foreach (DataColumn dc in dt.Columns)
            { 
                sb.Append(WriteCSV(dc.ColumnName.ToString()) + ",");
            }
            //this is for data display
            foreach (DataRow dr in dt.Rows)
            {

                foreach (DataColumn dc in dt.Columns)
                    sb.Append(WriteCSV(dr[dc.ColumnName].ToString()) + ",");
                sb.Remove(sb.Length - 1, 1);
                sb.AppendLine();
            }
            //if (File.Exists(@"C:\Users\16476\Documents\Test\vj.csv"))
            //{
            //    File.Delete(@"C:\Users\16476\Documents\Test\vj.csv");
            //}
            File.WriteAllText(@"C:\Users\16476\Documents\Test\vj.csv", sb.ToString());
            sendEMailThroughOUTLOOK();

        }
    

        public DataTable GetData()
        {

            SqlConnection sqlConn = new SqlConnection(@"Data Source=DESKTOP-C2JRKQS\SQLEXPRESS;Initial Catalog=Vijay;Integrated Security=True");
            sqlConn.Open();
            SqlDataAdapter daa = new SqlDataAdapter("SELECT  *  FROM Employee", sqlConn);
            using (DataTable dt = new DataTable())
            {
                daa.Fill(dt);
                return dt;

            }


        }


        public static string WriteCSV(string input)
        {
            try
            {
                if (input == null)
                    return string.Empty;

                bool containsQuote = false;
                bool containsComma = false;
                int len = input.Length;
                for (int i = 0; i < len && (containsComma == false || containsQuote == false); i++)
                {
                    char ch = input[i];
                    if (ch == '"')
                        containsQuote = true;
                    else if (ch == ',')
                        containsComma = true;
                }

                if (containsQuote && containsComma)
                    input = input.Replace("\"", "\"\"");

                if (containsComma)
                    return "\"" + input + "\"";
                else
                    return input;
            }
            catch
            {
                throw;
            }
        }
        public void sendEMailThroughOUTLOOK()
        {
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                oMsg.HTMLBody = "Hello vijay please find the document attached.";
                //Add an attachment.
                String sDisplayName = "MyAttachment";
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //now attached the file
                Outlook.Attachment oAttach = oMsg.Attachments.Add
                                             (@"C:\Users\16476\Documents\Test\vj.csv", iAttachType, iPosition, sDisplayName);
                //Subject line
                oMsg.Subject = "Monthly Report.";
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                Outlook.Recipient oRecip1 = (Outlook.Recipient)oRecips.Add("vinishajonnavittulal@gmail.com");
                Outlook.Recipient oRecip2 = (Outlook.Recipient)oRecips.Add("vijaybhaskar.lanka@gmail.com");
                oRecip1.Resolve();
                oRecip2.Resolve();
                // Send.
                oMsg.Send();
                // Clean up.
                oRecip1 = null;
                oRecip2 = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }//end of try block
            catch (Exception ex)
            {
            }//end of catch
        }//end of Email Method

    }


}

