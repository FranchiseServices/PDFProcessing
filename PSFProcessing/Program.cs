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
using System.Threading;
using System.Threading.Tasks;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.io;


namespace PDFProcessing
{
    class Program
    {
        static List<int> neededPages = new List<int>();
        static DataTable emailList = new DataTable();

        static void Main(string[] args)
        {
            Utilities.WriteLogFile("Process Began " + DateTime.Now);



            //we create the global look up table that will hold the files to be emailed
            emailList.Columns.Add(new DataColumn("pdfPath"));
            emailList.Columns.Add(new DataColumn("pdfFileName"));
            emailList.Columns.Add(new DataColumn("Owner_ID"));
            emailList.Columns.Add(new DataColumn("invoiceID"));
            emailList.Columns.Add(new DataColumn("brand"));

                SearchForPDFProcess();

                //now that we have cycled throgh, grabbed all the pdf, parsed and split the files, we can now email them
                PushEmail();

        }

        static void SearchForPDFProcess()
        {

            // search for the designated "New" path where files are deposited by accounting
            string PDF_FilePath_New = System.Configuration.ConfigurationManager.AppSettings["PDFPathNew"];
            string PDF_ArchivePath_New = System.Configuration.ConfigurationManager.AppSettings["PDFPathArchinve"];
            string PDF_ArchivePath_Original = System.Configuration.ConfigurationManager.AppSettings["PDFPathOriginal"];
            string PDF_Path_Archive_Mailed_Base = System.Configuration.ConfigurationManager.AppSettings["PDFPathMailed"];
            //string Center_ID = string.Empty;
            string pdfArchivePath = string.Empty;
            string newPDFFileName = string.Empty;
            string previousInvoiceID = string.Empty;
            string pdfArchiveOriginalFile = string.Empty;

            int OwnerIDidx = 0;
            int invoiceIDidx = 0;
            int stringLength = 0;


            string NewFileName = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString().Replace("\\", "-");
            string sYear = DateTime.Now.Year.ToString();
            string sMonth = DateTime.Now.ToString("MMMM");
            string ArchiveOriginal = System.Configuration.ConfigurationManager.AppSettings["PDFPathOriginal"];
            string pageText = string.Empty;
            string Owner_ID = string.Empty;
            string InvoiceID = string.Empty;


            string previousPageNumber = string.Empty;
            string previousPDF_ArchivePath = string.Empty;
            string previousNewFileName = string.Empty;
            string previousOwner_ID = string.Empty;
            string previous_InvoiceID = string.Empty;
            string previous_Brand = string.Empty;


            int numberofPages = 0;
            PdfReader reader = null;
            List<int> concateInvoices = new List<int>();
            string Brand = string.Empty;
            int idxUnderScore = 0;

            string pdfArchivePathFull = string.Empty;
            string pdfArchivePathBase = string.Empty;


            if (Directory.GetFiles(PDF_FilePath_New, "*.pdf").Length != 0)
            {

                foreach (string fi in Directory.GetFiles(PDF_FilePath_New, "*.pdf"))
                {
                    string baseFileName = "";

                    // log that files were found
                    Utilities.WriteLogFile(Environment.NewLine + "file found " + fi + Environment.NewLine);


                    int idxSlash = fi.LastIndexOf(@"\");
                    baseFileName = fi.Substring(idxSlash + 1, (fi.Length - 1) - idxSlash);


                    idxUnderScore = baseFileName.IndexOf("_");
                    Brand = baseFileName.Substring(0, idxUnderScore);


                    pdfArchiveOriginalFile = DateTime.Now.ToLongDateString() + "-" + baseFileName;


                    //here we test if the file has been created
                    if (!System.IO.File.Exists(PDF_ArchivePath_Original + "\\" + pdfArchiveOriginalFile))
                    {
                        File.Move(fi, (PDF_ArchivePath_Original + "\\" + pdfArchiveOriginalFile));
                        Utilities.WriteLogFile(Environment.NewLine + "pdf file created " + PDF_ArchivePath_Original + "\\" + pdfArchiveOriginalFile + Environment.NewLine);
                    }

                    //grab the newly found pdf file and load it into the reader to loop tuthrough the pages
                    reader = new PdfReader(PDF_ArchivePath_Original + "\\" + pdfArchiveOriginalFile);
                    //grab the number of pages for the purposes of the boundary of the FOR loop
                    numberofPages = reader.NumberOfPages;

                    try
                    {
                        for (int pageNumber = 1; pageNumber <= numberofPages; pageNumber++)
                        {


                            //here we grab the text from the page and look as certain text values                            
                            //ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
                            pageText = PdfTextExtractor.GetTextFromPage(reader, pageNumber);

                            if (pageText.Length < 1)
                            {
                                // maybe a binarry file and not readable, send email to dev with filename
                                string strfilepath = string.Empty;
                                strfilepath = PDF_ArchivePath_Original + "\\" + pdfArchiveOriginalFile;
                                String devEmail = ConfigurationManager.AppSettings["devEmail"];


                                Utilities.WriteLogFile(Environment.NewLine + "bad file, possibly binary pdf " + strfilepath + Environment.NewLine);
                                Utilities.SendEmailWithAttachment("bad file, possibly binary pdf " + strfilepath, "Error in reading  PDF on page " + pageNumber.ToString(), devEmail);

                                break;
                            }

                            //this grabs the string index of where is phrase "Owner:"... next to it is the franchise id
                            OwnerIDidx = pageText.IndexOf("TLIB");
                            //here we grad the string index of where the phrase "Invoice:"... next to this is the invoice number
                            invoiceIDidx = pageText.IndexOf("Invoice#:");
                            //this is just the length of the string
                            stringLength = pageText.Length;

                            //this is the invoice id as pulled from the string
                            InvoiceID = pageText.Substring(invoiceIDidx + 10, 7);
                            //here we grab the whole value of the ownerid
                            Owner_ID = pageText.Substring(OwnerIDidx, 9);
                            //we have to strip the letter "B" out


                            //This will be the path of the invoice pdf that will be emailed and archived
                            pdfArchivePath = PDF_Path_Archive_Mailed_Base + "\\" + Owner_ID + "\\" + sYear;

                            newPDFFileName = pdfArchivePath + "\\" + Owner_ID + "-" + sYear + "-" + sMonth + "-" + InvoiceID + ".pdf";


                            //here we look to see if the path exists, if not, create it
                            if (!System.IO.Directory.Exists(pdfArchivePath))
                            {
                                System.IO.Directory.CreateDirectory(pdfArchivePath);
                            }

                            //here we add the previous page number
                            if (pageNumber != 1)
                            {
                                neededPages.Add(pageNumber - 1);
                            }

                            // we are testing to see if it will become a new document
                            if (previousInvoiceID != InvoiceID)
                            {
                                // this will create a new PDF file
                                newPDFFileName = CreateNewPDF(pageNumber, PDF_ArchivePath_Original, pdfArchiveOriginalFile, Owner_ID, InvoiceID, Brand);

                                pdfArchivePathBase = System.Configuration.ConfigurationManager.AppSettings["PDFPathMailed"];
                                pdfArchivePathFull = pdfArchivePathBase + "\\" + previousOwner_ID + "\\" + sYear;

                                neededPages = new List<int>();

                                previousInvoiceID = InvoiceID;

                                //set the variables for the next round trip
                                previousPageNumber = pageNumber.ToString();
                                previousPDF_ArchivePath = pdfArchivePath;
                                previousNewFileName = newPDFFileName;
                                previousOwner_ID = Owner_ID;
                                previous_InvoiceID = InvoiceID;
                                previous_Brand = Brand;

                                neededPages.Add(pageNumber);

                            }
                            else
                            {
                                // make sure it is passing in the previous path for the pdf to add to
                                AddPageToPDF(pageNumber, PDF_ArchivePath_Original, pdfArchiveOriginalFile, previousPDF_ArchivePath, previousNewFileName, Owner_ID, InvoiceID);
                            }

                        }



                    }
                    catch (Exception ex)
                    {
                        Utilities.WriteLogFile(Environment.NewLine + ex.ToString() + Environment.NewLine);
                    }
                    finally
                    {
                        reader.Close();

                    }


                }
            }
            else
            {
                Utilities.WriteLogFile(Environment.NewLine + "no files present " + DateTime.Now.ToLongDateString() + Environment.NewLine);
            }


        }

        static string GetEmail(string Owner_ID)
        {
            string returnVal = string.Empty;
            string sproc = "get_email_address_for_owner_number";
            string connString = ConfigurationManager.AppSettings["connectionString"];

            try
            {

                SqlConnection conn = new SqlConnection(connString);
                conn.Open();
                SqlCommand cmd = new SqlCommand(sproc, conn);
                { cmd.CommandType = CommandType.StoredProcedure; };

                cmd.Parameters.AddWithValue("@owner_number", Owner_ID);

                returnVal = cmd.ExecuteScalar().ToString();

                conn.Close();
                conn.Dispose();
                cmd.Dispose();
            }
            catch (Exception ex)
            {
                Utilities.WriteLogFile(" GetEmail" + Environment.NewLine + ex.ToString() + Environment.NewLine);
                returnVal = string.Empty;
            }


            return returnVal;
        }

        static void EmailPDF(string pdfPath, string pdfFileName, string Owner_Number, string invoiceID, string Brand)
        {
            String EmailAddress = string.Empty;
           // String smtpServer = ConfigurationManager.AppSettings["smtpServer"];
            String devEmail = ConfigurationManager.AppSettings["devEmail"];

            String mailBody = string.Empty;
            String emailSubject = ConfigurationManager.AppSettings["EmailSubject"];


            //debugging code
            if (ConfigurationManager.AppSettings["debug"] == "yes")
            {
                EmailAddress = devEmail;
            }
            else
            {
                EmailAddress = GetEmail(Owner_Number);
            }


            if (EmailAddress.Trim().Length == 0)
            {
                Utilities.WriteLogFile(Environment.NewLine + "Email was blank for OwnerNumber = " + Owner_Number  + Environment.NewLine);
                Utilities.SendEmailWithAttachment(Environment.NewLine + "Email was blank for OwnerNumber = " + Owner_Number + Environment.NewLine, " Error no email address",devEmail);
                return;
            }



            List<string> emailAddresses = new List<string>();

            if (EmailAddress.Contains(","))
            {
                //split string to list
                emailAddresses = EmailAddress.Split(',').ToList();
            }
            else
            {
                //just a single address was retirned and just add it to the array
                emailAddresses.Add(EmailAddress);
            }


            try
            {
                foreach (string email in emailAddresses)
                {
                    //test that the email address is of a valid structure
                    if (Utilities.IsValidEmail(email.Trim()) != true)
                    {
                        Utilities.SendEmailWithAttachment(email.Trim() + " is a bad email address " + Environment.NewLine + "Center ID " + Owner_Number + " invoice " + invoiceID, "Bad Email Address", devEmail);
                        break;
                    }


                    mailBody = "<div style='padding: 15px;text-align:left;display:block;overflow:auto;font-family:Arial;font-size:12px;margin:auto;'>"; ;
                    mailBody += "<div style='padding: 15px;'>";


                    if (Brand.ToLower().Contains("royal"))
                    {
                        mailBody += System.Configuration.ConfigurationManager.AppSettings["RoyaltyLetter"];
                    }


                    if (Brand.ToLower().Contains("kase") )
                    {
                        mailBody += System.Configuration.ConfigurationManager.AppSettings["KaseyaLetter"];
                    }

                    if ( Brand.ToLower().Contains("auto"))
                    {
                        mailBody += System.Configuration.ConfigurationManager.AppSettings["AutoTaskLetter"];
                    }



                    mailBody += "</div>";
                    mailBody += "</ br></ br></ br></ br></ br></ br></ br></ br></ br></ br></ br></ br></ br></ br></ br></ br></ br></ br><div style='position: fixed; bottom: 0;font-weight:bold;padding-top:50px;'>";
                    mailBody += ConfigurationManager.AppSettings["FooterLetter"];
                    mailBody += "</div></div>";


                    String fileName = pdfPath + pdfFileName;
                    String mailSubject = emailSubject + " billing for " + Owner_Number + " Inv# " + invoiceID;


                    Utilities.SendEmailWithAttachment(mailBody.Trim(),mailSubject, email, fileName);

                    LogEmail(Owner_Number, pdfFileName, pdfPath, invoiceID, email.ToString());

                }
            }
            catch (Exception ex)
            {
                Utilities.WriteLogFile("Error in Mail Function" + Environment.NewLine + ex.ToString() + Environment.NewLine);
                Utilities.SendEmailWithAttachment("Error in Mail Function" + Environment.NewLine + ex.ToString() + Environment.NewLine, "EmailPDF",devEmail);
            }

        }

        static void PushEmail()
        {
            string pdfPath = string.Empty;
            string pdfFileName = string.Empty;
            string Owner_ID = string.Empty;
            string invoiceID = string.Empty;
            string Brand = string.Empty;
            int idxDash = 0;


            try
            {
                foreach (DataRow row in emailList.Rows)
                {

                    pdfPath = row["pdfPath"].ToString();
                    pdfFileName = row["pdfFileName"].ToString();
                    Owner_ID = row["Owner_ID"].ToString();
                    invoiceID = row["invoiceID"].ToString();

                    //grab the brand from the filename here brand - 
                    idxDash = pdfFileName.IndexOf("-");

                    if (idxDash >= 0)
                    {
                        Brand = pdfFileName.Substring(0, idxDash);
                    }


                    //validate that it has not been emailed before
                    if (HasInvoiceBeenEmailed(invoiceID, Owner_ID) == false)
                    {
                        EmailPDF(pdfPath, pdfFileName, Owner_ID, invoiceID, Brand);
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.WriteLogFile("Push Email " + ex.ToString() + Environment.NewLine);
            }
        }

        static void LogEmail(string Owner_ID, string pdf_file_name, string pdf_path, string invoice_number, string email_address)
        {

            try
            {
                string sproc = "add_to_transmittal_invoice_email_log";
                string connString = ConfigurationManager.AppSettings["connectionString"];

                SqlConnection conn = new SqlConnection(connString);
                SqlCommand cmd = new SqlCommand(sproc, conn);
                { cmd.CommandType = CommandType.StoredProcedure; };

                cmd.Parameters.AddWithValue("@owner_id", Owner_ID);
                cmd.Parameters.AddWithValue("@pdf_file_name", pdf_file_name);
                cmd.Parameters.AddWithValue("@pdf_path", pdf_path);
                cmd.Parameters.AddWithValue("@invoice_number", invoice_number);
                cmd.Parameters.AddWithValue("@email_address", email_address);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Utilities.WriteLogFile("Error in LogEmail" + Environment.NewLine + ex.ToString() + Environment.NewLine);
            }
        }

        static bool HasInvoiceBeenEmailed(string Invoice_Number, string Owner_ID)
        {
            bool returnValue = false;
            SqlDataReader dr = null;
            string outputValue = string.Empty;

            string sproc = "has_invoice_been_emailed";
            string connString = ConfigurationManager.AppSettings["connectionString"];

            SqlConnection conn = new SqlConnection(connString);
            conn.Open();
            SqlCommand cmd = new SqlCommand(sproc, conn);
            { cmd.CommandType = CommandType.StoredProcedure; };

            cmd.Parameters.AddWithValue("@invoice_number", Invoice_Number);
            cmd.Parameters.AddWithValue("@owner_id", Owner_ID);

            dr = cmd.ExecuteReader();
            returnValue = dr.HasRows;

            // if (Utilities.IsNumeric(outputValue) == true) { returnValue = true; };
            if (returnValue == true)
            {
                //record that the invoice had been emailed
                Utilities.WriteLogFile(" Invoice had been emailed for " + Owner_ID + " invoice# " + Invoice_Number + Environment.NewLine);
            }



            return returnValue;
        }

        static string CreateNewPDF(int pageNumber, string pdfArchivePathOriginal, string pdfArchiveOriginalFile, string Owner_ID, string invoiceID, string Brand)
        {

            string pdfArchiveMailedPath = System.Configuration.ConfigurationManager.AppSettings["PDFPathMailed"];

            string pdfArchivedMailFileName = string.Empty;
            string sYear = DateTime.Now.Year.ToString();
            string sMonth = DateTime.Now.ToString("MMMM");

            string newPdfArchiveMailedPath = string.Empty;
            string newPdfArchiveMailedFileName = string.Empty;
            string newPDFPathAndFileName = string.Empty;


            // declare pdf object variable
            PdfReader reader = null;
            Document document = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            newPdfArchiveMailedPath = pdfArchiveMailedPath + Owner_ID + "\\" + sYear + "\\";


            try
            {


                //here we look to see if the path exists, if not, create it
                if (!System.IO.Directory.Exists(newPdfArchiveMailedPath))
                {
                    System.IO.Directory.CreateDirectory(newPdfArchiveMailedPath);
                }

                newPdfArchiveMailedFileName = Brand + "-" + Owner_ID + "-" + sYear + "-" + sMonth + "-" + invoiceID + ".pdf";

                //this will be the path and file name for the new pdf document for the specific center_id, invoice number
                newPDFPathAndFileName = newPdfArchiveMailedPath + "\\" + newPdfArchiveMailedFileName;


                if (File.Exists(newPDFPathAndFileName))
                {
                    //file already exists bail out
                    if (document != null) { document.Dispose(); };
                    if (reader != null) { reader.Dispose(); };

                    Utilities.WriteLogFile(" File Exists for Create New PDF " + newPDFPathAndFileName + Environment.NewLine);
                    return string.Empty;
                }

                //grab the base file with all the invoices
                string filepath = pdfArchivePathOriginal + "\\" + pdfArchiveOriginalFile;
                reader = new PdfReader(filepath);




                //grab the exact page number from within the PDF
                document = new Document(reader.GetPageSizeWithRotation(pageNumber));

                //here we create the new pdf file name
                pdfCopyProvider = new PdfCopy(document, new FileStream(newPDFPathAndFileName, FileMode.Create));

                document.Open();

                //look at old reader and grab the page number
                importedPage = pdfCopyProvider.GetImportedPage(reader, pageNumber);
                //add the single page and add to the new pdf
                pdfCopyProvider.AddPage(importedPage);


                //we add the info of the location of the pdf so it can be emailed
                DataRow Row = emailList.NewRow();

                Row["pdfPath"] = newPdfArchiveMailedPath;
                Row["pdfFileName"] = newPdfArchiveMailedFileName;
                Row["Owner_ID"] = Owner_ID;
                Row["invoiceID"] = invoiceID;
                emailList.Rows.Add(Row);


                //clean up
                document.Close();
                reader.Close();

                //return complete path + filename of the exact file
                return newPdfArchiveMailedFileName;


            }
            catch (Exception ex)
            {
                Utilities.WriteLogFile(" Error in CreateNewPDF " + ex.ToString() + Environment.NewLine);
                return string.Empty;
            }
            finally
            {
                if (document != null) { document = null; };
                if (reader != null) { reader = null; };
            }


        }

        static string AddPageToPDF(int pageNumber, string pdfArchivePathOriginal, string pdfArchiveOriginalFileName, string pdfArchiveMailedPath, string pdfArchivedMailFileName, string Owner_ID, string invoiceID)
        {
            string returnValue = string.Empty;

            try
            {
                
                PdfReader reader = null;

                if (File.Exists(pdfArchiveMailedPath + "\\" + pdfArchivedMailFileName) == true)
                {

                    neededPages.Add(pageNumber);

                    var destinationDocumentStream = new FileStream(pdfArchiveMailedPath + "\\" + pdfArchivedMailFileName, FileMode.Open);
                    var pdfConcat = new PdfConcatenate(destinationDocumentStream);

                    reader = new PdfReader(pdfArchivePathOriginal + "\\" + pdfArchiveOriginalFileName);

                    reader.SelectPages(neededPages);
                    pdfConcat.AddPages(reader);

                    reader.Close();
                    pdfConcat.Close();

                }
                returnValue = pdfArchivedMailFileName;

            }
            catch (Exception ex)
            {
                Utilities.WriteLogFile("AddPageToPDF " + Environment.NewLine + ex.ToString() + Environment.NewLine);
                returnValue = "error";
            }


            return returnValue;
        }

    }
}
