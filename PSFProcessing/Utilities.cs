using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFProcessing
{
    class Utilities
    {
        public static bool IsValidEmail(string email)
        {
            try
            {
                MailAddress addr = new MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        public static bool IsNumeric(String numberString)
        {
            double output;
            return double.TryParse(numberString, out output);
        }

        public static void WriteLogFile(string logText)
        {

            //This will give the log file name of 01-01-1970.log for example.
            String logDate = DateTime.Now.ToShortDateString().Replace(@"/", "-");
            String logName = logDate + ".log";
            String logPath = string.Empty;

            //Grab the log file path from the web config.
            logPath = System.Configuration.ConfigurationManager.AppSettings["LogFilePath"].ToString();
            
                      
            //Grab the value for the maximum days for the age of the log files before deleting.
            int maxDays = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["LogMaxDays"].ToString());
            //Delete the old log files.
            DeleteOldLogFiles(logPath, maxDays);


            DirectoryInfo logDir = new DirectoryInfo(logPath);

            // Create a writer and open the file:
            StreamWriter w;

            //Here we check to see if the file exists, then we can log an event, otherwise, create the log file.
            if (System.IO.File.Exists(logPath + @"\" + logName) == true)
            {
                w = System.IO.File.AppendText(logPath + @"\" + logName);
                w.WriteLine(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " - " + logText);
                w.Flush();
            }
            else
            {
                w = System.IO.File.CreateText(logPath + @"\" + logName);
                w.WriteLine(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " - " + logText);
                w.Flush();
            }



            // Close the stream and destroy the stream object.
            w.Close();
            w.Dispose();
        }

        public static void DeleteOldLogFiles(string path, int maxDays)
        {

            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo file = null;

            if (dir.Exists == false) { return; };

            try
            {
                //Cycle through all the files in the selected directory.
                foreach (object tempFile in dir.GetFiles())
                {
                    file = (FileInfo)(tempFile);
                    //Delete file if it is over 30 days old.
                    if (file.CreationTime.AddDays(maxDays) < DateTime.Now) { file.Delete(); };
                    if (file != null) { file = null; };
                }
            }
            catch (Exception Ex)
            {
                WriteLogFile(Ex.Message + " " + Ex.StackTrace);
            }
            finally
            {
                if (file != null) { file = null; };
                if (dir != null) { dir = null; };
            }


        }

        public static void SendEmailWithAttachment(string emailMessage, string emailSubject, string emailAddressTo, string attachmentPath = "")
        {
            String smtpServer = ConfigurationManager.AppSettings["smtpServer"];
            String devEmail = ConfigurationManager.AppSettings["devEmail"];
            string mailUser = ConfigurationManager.AppSettings["username"];
            string mailPassword = ConfigurationManager.AppSettings["password"];


            string MailBody = string.Empty;
            try
            {
                MailMessage message = new MailMessage();
                message.IsBodyHtml = true;
                MailAddress outGoingEmailAddressFromAdmin = new MailAddress("no-reply@franserv.com");
                message.From = outGoingEmailAddressFromAdmin;

                if (ConfigurationManager.AppSettings["debug"] == "yes")
                {
                    emailAddressTo = devEmail;
                }

                MailAddress mailAddressTo = new MailAddress(emailAddressTo);

                message.To.Add(mailAddressTo);
                message.Bcc.Add(new MailAddress(devEmail));

                message.Subject = emailSubject;

                MailBody = "<div style='padding: 15px;text-align:left;'>" + emailMessage + "</div>";
                message.Body = MailBody;

                if (attachmentPath.Length > 0)
                {
                    message.Attachments.Add(new Attachment(attachmentPath));
                }

                SmtpClient smtpClient = new SmtpClient();
                SmtpClient client = smtpClient;
                client.Credentials = new System.Net.NetworkCredential(mailUser, mailPassword);
                client.Port = 25;
                client.Host = smtpServer;
                client.EnableSsl = true;
                client.Send(message);

                client.Dispose();
            }
            catch (Exception Ex)
            {
                WriteLogFile(Ex.Message + " " + Ex.StackTrace);
            }

        }

    }
}
