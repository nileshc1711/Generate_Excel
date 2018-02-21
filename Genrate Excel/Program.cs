using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;
using System.Drawing;
using Renci.SshNet;
using System.IO;
using System.Configuration;
using log4net;
using System.Threading;
using System.Net.Mail;

namespace Genrate_Excel
{
    class Program
    {
        static void Main(string[] args)
        {

            Logging lng = new Logging();

            lng.log.Info("Starting Application");
            Excel exl = new Excel();
            exl.GenrateExcel();

            lng.log.Info("Excel Genration has been completed");

            lng.log.Info("starting ftp connection to upload file");

            FTP f = new FTP(ConfigurationManager.AppSettings["host"].ToString(), ConfigurationManager.AppSettings["username"].ToString(), ConfigurationManager.AppSettings["password"].ToString());

            lng.log.Info("uploading file");

            f.UploadFile("D:\\temp.xlsx", "temp.xlsx");

            Console.WriteLine("Excel has been genrated and uploaded to FTP");

            Email em = new Email();
            em.sendemail();

            lng.log.Info("closing application");
        }
    }
    public class Excel
    {
        Logging lng = new Logging();
        public void GenrateExcel()
        {
            Logging lng = new Logging();
            lng.log.Info("Starting for Genrating Excel Function");
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);

                lng.log.Info("creating worksheet of excel");
                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "Excel";

                worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
                //worKsheeT.Cells[1, 1] = "Excel";
                worKsheeT.Cells.Font.Size = 11;

                lng.log.Info("Adding data from export excel function and looping into the function");
                int rowcount = 2;

                foreach (DataRow datarow in ExportToExcel().Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= ExportToExcel().Columns.Count; i++)
                    {

                        if (rowcount == 3)
                        {
                            worKsheeT.Cells[2, i] = ExportToExcel().Columns[i - 1].ColumnName;
                            worKsheeT.Cells.Font.Color = System.Drawing.Color.Black;

                        }

                        worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

                        if (rowcount > 3)
                        {
                            if (i == ExportToExcel().Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, ExportToExcel().Columns.Count]];
                                }

                            }
                        }

                    }

                }


                lng.log.Info("Setting cellramnge");

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, ExportToExcel().Columns.Count]];
                celLrangE.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                lng.log.Info("Setting border");


                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, ExportToExcel().Columns.Count]];
                string path = ConfigurationManager.AppSettings["path"].ToString();
                lng.log.Info("saving excel to path");
                worKbooK.SaveAs(path);

                worKbooK.Close();
                lng.log.Info("closing excel");
                excel.Quit();
                lng.log.Info("Quiting from excel");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                lng.log.Fatal(ex.Message.ToString());
                lng.log.Fatal(ex.StackTrace.ToString());

            }
            finally
            {
                lng.log.Info("setting null to excel");
                worKsheeT = null;
                celLrangE = null;
                worKbooK = null;
            }
        }

        public  System.Data.DataTable ExportToExcel()
        {

            lng.log.Info("Export to excel function");
            System.Data.DataTable table = new System.Data.DataTable();

            lng.log.Info("adding column to table");
            table.Columns.Add("Reverse Charge Transaction (Y or N)", typeof(string));
            table.Columns.Add("Posting charges for", typeof(string));
            table.Columns.Add("Enc. ID (not for Treat. Series)", typeof(string));
            table.Columns.Add("ECD ID (only for Treat. Series)", typeof(string));
            table.Columns.Add("Service Provider Service ID", typeof(string));
            table.Columns.Add("Service Date", typeof(string));
            table.Columns.Add("Service Time", typeof(string));
            table.Columns.Add("Service Time DST/ST", typeof(string));
            table.Columns.Add("Stop Date", typeof(string));
            table.Columns.Add("Stop Time", typeof(string));
            table.Columns.Add("Stop Time DST/ST", typeof(string));
            table.Columns.Add("Quantity", typeof(string));
            table.Columns.Add("Extended Price", typeof(string));
            table.Columns.Add("Duration (in minutes)", typeof(string));
            table.Columns.Add("Performing HP ID", typeof(string));
            table.Columns.Add("Performing HP ID Issuer", typeof(string));
            table.Columns.Add("Ordering HP ID", typeof(string));
            table.Columns.Add("Ordering HP ID Issuer", typeof(string));
            table.Columns.Add("Referring HP ID", typeof(string));
            table.Columns.Add("Referring HP ID Issuer", typeof(string));
            table.Columns.Add("Supervising HP ID", typeof(string));
            table.Columns.Add("Supervising HP ID Issuer", typeof(string));
            table.Columns.Add("Override Proc. Code", typeof(string));
            table.Columns.Add("Proc. Mod. (1)", typeof(string));
            table.Columns.Add("Proc. Mod. (2)", typeof(string));
            table.Columns.Add("Proc. Mod. (3)", typeof(string));
            table.Columns.Add("Proc. Mod. (4)", typeof(string));
            table.Columns.Add("Override Rev. Code", typeof(string));
            table.Columns.Add("Override Charge Amt.", typeof(string));
            table.Columns.Add("Override Service Name", typeof(string));
            table.Columns.Add("Diag. Code (1)", typeof(string));
            table.Columns.Add("Diag. Code (2)", typeof(string));
            table.Columns.Add("Diag. Code (3)", typeof(string));
            table.Columns.Add("Diag. Code (4)", typeof(string));
            table.Columns.Add("Diag. Code (5)", typeof(string));
            table.Columns.Add("Diag. Code (6)", typeof(string));
            table.Columns.Add("Diag. Code (7)", typeof(string));
            table.Columns.Add("Diag. Code (8)", typeof(string));
            table.Columns.Add("Dose Quantity", typeof(string));
            table.Columns.Add("Referral Number", typeof(string));
            table.Columns.Add("Authorization Number", typeof(string));
            table.Columns.Add("Cost", typeof(string));
            table.Columns.Add("ABN Status Date", typeof(string));
            table.Columns.Add("ABN Status Code", typeof(string));
            table.Columns.Add("National Drug Code", typeof(string));
            table.Columns.Add("Item Number", typeof(string));
            table.Columns.Add("Model Number", typeof(string));
            table.Columns.Add("Taxonomy Code", typeof(string));
            table.Columns.Add("Clinic Code", typeof(string));
            table.Columns.Add("Tooth Designation", typeof(string));
            table.Columns.Add("Dent. Srfc. Code1", typeof(string));
            table.Columns.Add("Dent. Srfc. Code2", typeof(string));
            table.Columns.Add("Dent. Srfc. Code3", typeof(string));
            table.Columns.Add("Dent. Srfc. Code4", typeof(string));
            table.Columns.Add("Dent. Srfc. Code5", typeof(string));
            table.Columns.Add("Tooth Status Code", typeof(string));
            table.Columns.Add("Oral Cavity Code1", typeof(string));
            table.Columns.Add("Oral Cavity Code2", typeof(string));
            table.Columns.Add("Oral Cavity Code3", typeof(string));
            table.Columns.Add("Oral Cavity Code4", typeof(string));
            table.Columns.Add("Oral Cavity Code5", typeof(string));
            table.Columns.Add("Placement Status Code", typeof(string));
            table.Columns.Add("Prior Placement Date", typeof(string));
            table.Columns.Add("Podiatry Last PCP Visit Date", typeof(string));
            table.Columns.Add("Hearing And Vision Prescription Date", typeof(string));
            table.Columns.Add("Vision Category Code", typeof(string));
            table.Columns.Add("Vision Certification Condition Indicator", typeof(string));
            table.Columns.Add("Vision Condition Indicator Code1", typeof(string));
            table.Columns.Add("Vision Condition Indicator Code2", typeof(string));
            table.Columns.Add("Vision Condition Indicator Code3", typeof(string));
            table.Columns.Add("Vision Condition Indicator Code4", typeof(string));
            table.Columns.Add("Vision Condition Indicator Code5", typeof(string));
            table.Columns.Add("Dme Certificate of Medical Necessity Transmission Code", typeof(string));
            table.Columns.Add("Dme Certification Type Code", typeof(string));
            table.Columns.Add("Dme Duration(Months)", typeof(string));
            table.Columns.Add("Dme Certification Revision Date", typeof(string));
            table.Columns.Add("Dme initial Certification Date", typeof(string));
            table.Columns.Add("Dme Last Certification Date", typeof(string));
            table.Columns.Add("Dme Length of Medical Necessity Days", typeof(string));
            table.Columns.Add("Dme Rental Price", typeof(string));
            table.Columns.Add("Dme Purchase Price", typeof(string));
            table.Columns.Add("Dme Frequency Code", typeof(string));
            table.Columns.Add("Dme Certification Condition Indicator", typeof(string));
            table.Columns.Add("Dme Condition Indicator Code1", typeof(string));
            table.Columns.Add("Dme Condition Indicator Code2", typeof(string));
            table.Columns.Add("Special Processing Code1", typeof(string));
            table.Columns.Add("Special Processing Code2", typeof(string));
            table.Columns.Add("Special Processing Code3", typeof(string));
            table.Columns.Add("Special Processing Code4", typeof(string));
            table.Columns.Add("Special Processing Code5", typeof(string));
            table.Columns.Add("Special Processing Code6", typeof(string));
            table.Columns.Add("Investigational Device Exempt No", typeof(string));
            table.Columns.Add("Place Of Service Override", typeof(string));
            table.Columns.Add("Procedure Code Description Override", typeof(string));
            table.Columns.Add("Enc. Provider", typeof(string));
            table.Columns.Add("Enc. Location", typeof(string));
            table.Columns.Add("Enc. Matching Strategy Code", typeof(string));
            table.Columns.Add("National Drug Code Quantity", typeof(string));
            table.Columns.Add("RX Number", typeof(string));
            table.Columns.Add("Unit Of Measure Code", typeof(string));
            table.Columns.Add("Auto Charge Rule Text", typeof(string));
            table.Columns.Add("Auto Charge Rule Type Code", typeof(string));
            table.Columns.Add("Service Price Rule Text", typeof(string));
            table.Columns.Add("Error Message", typeof(string));
            lng.log.Info("added column to table");


            lng.log.Info("adding rows to table");
            table.Rows.Add(1, "test", "M", 78, 59, 72, 95, 83, 77);



            lng.log.Info("adding rows to table");
            table.Rows.Add("Reverse Charge Transaction (Y or N)", "Posting charges for", "Enc. ID (not for Treat. Series)", "ECD ID (only for Treat. Series)", "Service Provider Service ID", "Service Date", "Service Time", "Service Time DST/ST", "Stop Date", "Stop Time", "Stop Time DST/ST", "Quantity", "Extended Price", "Duration (in minutes)", "Performing HP ID", "Performing HP ID Issuer", "Ordering HP ID", "Ordering HP ID Issuer", "Referring HP ID", "Referring HP ID Issuer", "Supervising HP ID", "Supervising HP ID Issuer", "Override Proc. Code", "Proc. Mod. (1)", "Proc. Mod. (2)", "Proc. Mod. (3)", "Proc. Mod. (4)", "Override Rev. Code", "Override Charge Amt.", "Override Service Name", "Diag. Code (1)", "Diag. Code (2)", "Diag. Code (3)", "Diag. Code (4)", "Diag. Code (5)", "Diag. Code (6)", "Diag. Code (7)", "Diag. Code (8)", "Dose Quantity", "Referral Number", "Authorization Number", "Cost", "ABN Status Date", "ABN Status Code", "National Drug Code", "Item Number", "Model Number", "Taxonomy Code", "Clinic Code", "Tooth Designation", "Dent. Srfc. Code1", "Dent. Srfc. Code2", "Dent. Srfc. Code3", "Dent. Srfc. Code4", "Dent. Srfc. Code5", "Tooth Status Code", "Oral Cavity Code1", "Oral Cavity Code2", "Oral Cavity Code3", "Oral Cavity Code4", "Oral Cavity Code5", "Placement Status Code", "Prior Placement Date", "Podiatry Last PCP Visit Date", "Hearing And Vision Prescription Date", "Vision Category Code", "Vision Certification Condition Indicator", "Vision Condition Indicator Code1", "Vision Condition Indicator Code2", "Vision Condition Indicator Code3", "Vision Condition Indicator Code4", "Vision Condition Indicator Code5", "Dme Certificate of Medical Necessity Transmission Code", "Dme Certification Type Code", "Dme Duration(Months)", "Dme Certification Revision Date", "Dme initial Certification Date", "Dme Last Certification Date", "Dme Length of Medical Necessity Days", "Dme Rental Price", "Dme Purchase Price", "Dme Frequency Code", "Dme Certification Condition Indicator", "Dme Condition Indicator Code1", "Dme Condition Indicator Code2", "Special Processing Code1", "Special Processing Code2", "Special Processing Code3", "Special Processing Code4", "Special Processing Code5", "Special Processing Code6", "Investigational Device Exempt No", "Place Of Service Override", "Procedure Code Description Override", "Enc. Provider", "Enc. Location", "Enc. Matching Strategy Code", "National Drug Code Quantity", "RX Number", "Unit Of Measure Code", "Auto Charge Rule Text", "Auto Charge Rule Type Code", "Service Price Rule Text", "Error Message");

            lng.log.Info("adding rows to table");
            table.Rows.Add(1, "test", "M", 78, 59, 72, 95, 83, 77);



            return table;
        }
    }
    public class FTP
    {


        Logging lng = new Logging();
        public string Host { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public FTP(string host, string username, string password)
        {
            lng.log.Info("assing value to host,username ,password");
            Host = host;
            Username = username;
            Password = password;
        }
        public void UploadFile(string filepath, string filename)
        {
            lng.log.Info("creating connection to ftp");
            var connectionInfo = new ConnectionInfo(Host,
                                                Username,
                                                new PasswordAuthenticationMethod(Username, Password),
                                                new PrivateKeyAuthenticationMethod("rsa.key"));
            using (var client = new SftpClient(connectionInfo))
            {
                client.Connect();
                lng.log.Info("connected to ftp");

                lng.log.Info("getting file information for file upload");
                FileInfo f = new FileInfo(filepath);
                string uploadfile = f.FullName;

                lng.log.Info("checking whether connected to ftp or not");
                if (client.IsConnected)
                {
                    lng.log.Info("connected to ftp");
                }
                else
                {
                    lng.log.Info("not connected to ftp so connecting again");
                    client.Connect();
                }
                var fileStream = new FileStream(uploadfile, FileMode.Open);
                if (fileStream != null)
                {
                    lng.log.Info("file is not null so procedding for file upload");
                }
                client.BufferSize = 4 * 1024;
                client.UploadFile(fileStream, f.Name, null);
                Console.WriteLine("File has been uploaded");
                lng.log.Info("File has been uploaded");
                client.Disconnect();

                client.Dispose();
                lng.log.Info("disconnecting from ftp");
            }
        }
        public void DownloadFile(string remotefilepath)
        {
            lng.log.Info("creating connection to ftp");
            var connectionInfo = new ConnectionInfo(Host,
                                   Username,
                                   new PasswordAuthenticationMethod(Username, Password),
                                   new PrivateKeyAuthenticationMethod("rsa.key"));

            using (var client = new SftpClient(connectionInfo))
            {

                client.Connect();
                lng.log.Info("Connected to ftp");
                FileInfo f = new FileInfo(remotefilepath);
                string uploadfile = f.FullName;
                lng.log.Info("getting remote file path information");

                if (client.IsConnected)
                {
                    lng.log.Info("Connected to ftp");
                    Console.WriteLine("CONNECTED to ftp server");
                }
                else
                {
                    client.Connect();
                    lng.log.Info("Not Connected to ftp, so connecting again");
                }
                var fileStream = new FileStream(uploadfile, FileMode.Create);
                lng.log.Info("creating file");
                if (fileStream != null)
                {
                    lng.log.Info("checking if file is null");
                }
                else
                {
                    lng.log.Warn("File is null");
                    Console.WriteLine("File is null");
                }

                if (client.Exists(remotefilepath))
                {
                    client.DownloadFile(remotefilepath, fileStream);
                    Console.WriteLine("File has been downloaded");
                    lng.log.Info("File has been downloaded");
                }
                else
                {
                    Console.WriteLine("File not found on FTP Server to Download");
                    lng.log.Warn("File not found on FTP Server to Download");
                }

            }

        }
    }
    public class Email
    {
        Logging lng = new Logging();
        public void sendemail()
        {
            try
            {
                lng.log.Info("send email function");
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["emailsmtp"].ToString());
                lng.log.Info("smtp settings");
                mail.From = new MailAddress("appmail@medusind.com");
                mail.To.Add(ConfigurationManager.AppSettings["emailto"].ToString());
                mail.Subject = "Notification";
                mail.Body = "body  of eami;";
                lng.log.Info("email  body and subject");
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["emailusername"].ToString(), ConfigurationManager.AppSettings["emailpassword"].ToString());
                lng.log.Info("smtp id , password");
                SmtpServer.EnableSsl = true;
                lng.log.Info("enable ssl");
                SmtpServer.Send(mail);
                lng.log.Info("sendign email");
            }
            catch (Exception ex)
            {
                lng.log.Fatal(ex.ToString());
                lng.log.Fatal(ex.StackTrace.ToString());

            }

        }
    }
    class Logging
    {
        public  readonly log4net.ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    }
}

