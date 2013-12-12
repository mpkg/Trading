using System;
using System.Net;
using System.ServiceModel.Description;
using Microsoft.Crm.Services.Utility;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.IO;
using System.Text;
using org.pdfbox.pdmodel;
using org.pdfbox.util;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Data;

namespace CRMOnlineCRUD
{
    internal class Program
    {
        private static IOrganizationService _service;

        static string filename = string.Empty;
        static int attachmentcount = 0;
        static StringBuilder packlistfiles = new StringBuilder();
        //static string connectionstring = Properties.Settings.Default.DBConnectionString;
        static string excelConnStr;
        static string sender;
        static int processedcount = 0, unprocessedcount = 0, filesprocessed = 0;
        static StringBuilder emailbody = new StringBuilder();
        static string emaildate = string.Empty;

        private static void Main(string[] args)
        {
            CreateService();
            ProcessBhaskarPacklists();
            Console.Read();

        }

        static private void ProcessBhaskarPacklists()
        {
            string
                    sno = string.Empty,
                    sortno = string.Empty,
                    rollno = string.Empty,
                    shade = string.Empty,
                    grade = string.Empty,
                    invoiceno = string.Empty,
                    dispatchno = string.Empty,
                    lrno = string.Empty,
                    partyname = string.Empty,
                    soldto = string.Empty,
                    item = string.Empty,
                    dispatchdate = string.Empty,
                    transporter = string.Empty,
                    truckno = string.Empty;

            decimal mtrs = decimal.Zero;
            decimal netweight = decimal.Zero;
            decimal grossweight = decimal.Zero;
            int pieces = 1;
            int totalrolls = 0;
            emailbody.AppendLine();
            string[] split = new string[] { "\r\n" };
            int recordno, errorflag;

            DataTable dtPacklist = new DataTable();
            dtPacklist.TableName = "Packlist";
            
            #region read invoice pdf
            /*
            string[] invfiles = Directory.GetFiles(Properties.Settings.Default.invlocation);
            
            for (int l = 0; l < invfiles.Length; l++)
            {
                string invfilename = Path.GetFileName(invfiles[l]);
                char[] invfilenamechars = invfilename.ToCharArray();
                int invoiceparse = 0, lrparse = 0;
                PDDocument inv = PDDocument.load(invfiles[l]);
                PDFTextStripper invstripper = new PDFTextStripper();
                int a;
                string[] invcontents = invstripper.getText(inv).Split(split, StringSplitOptions.None);                
            }
                /*
                for (a = 0; a < invcontents.Length; a++)
                {
                    if (invcontents[a].ToLower().Equals("l.r. no"))
                        break;
                }
                invoiceno = invcontents[a + 1];

                lrno = invcontents[a + 6];

                transporter = invcontents[a + 8];

                truckno = invcontents[a + 9];

                if (truckno.ToLower().StartsWith("inv"))
                    truckno = "";

                if (!Int32.TryParse(invoiceno, out invoiceparse)
                    || !Int32.TryParse(lrno, out lrparse)
                    || (!truckno.ToLower().StartsWith("mp") && !truckno.ToLower().StartsWith("mh") && truckno.Length != 0)
                    || !transporter.ToLower().StartsWith("st"))
                    continue;
               
            cmd.CommandText = @"update PackingList set LRNo = ''";
                                // "' WHERE InvoiceNo = '" + invoiceno + "'";
            cmd.ExecuteNonQuery();
            cmd.CommandText = @"update PackingList set Transporter = ''";
                                //"' WHERE InvoiceNo = '" + invoiceno + "'";
            cmd.ExecuteNonQuery();
            cmd.CommandText = @"update PackingList set TruckNo = ''";
                                //"' WHERE InvoiceNo = '" + invoiceno + "'";
            cmd.ExecuteNonQuery();
                //File.Move(invfiles[l], Properties.Settings.Default.processedfilelocation + Path.GetFileName(invfiles[l]));
            */
            #endregion

            string[] files = Directory.GetFiles(Properties.Settings.Default.BhaskarPacklistSavePath);
            for (int j = 0; j < files.Length; j++)
            {
                try
                {
                    string filename = Path.GetFileName(files[j]);
                    Console.WriteLine("Processing File: " + filename);
                    int listcounter = 0;
                    errorflag = 0;
                    dtPacklist.Reset();

                    if (files[j].ToLower().EndsWith("pdf"))
                    {
                        try
                        {
                            string[] contents = PDFToTextPDFBox(files[j], split, StringSplitOptions.None);

                            dtPacklist.Columns.Add("F1", typeof(string));
                            dtPacklist.Columns.Add("F2", typeof(string));
                            dtPacklist.Columns.Add("F3", typeof(string));
                            dtPacklist.Columns.Add("F4", typeof(string));
                            dtPacklist.Columns.Add("F5", typeof(string));
                            dtPacklist.Columns.Add("F6", typeof(string));
                            dtPacklist.Columns.Add("F7", typeof(string));
                            dtPacklist.Columns.Add("F8", typeof(string));
                            dtPacklist.Columns.Add("F9", typeof(string));
                            dtPacklist.Columns.Add("F10", typeof(string));
                            dtPacklist.Columns.Add("F11", typeof(string));
                            dtPacklist.Columns.Add("F12", typeof(string));
                            dtPacklist.Columns.Add("F13", typeof(decimal));
                            dtPacklist.Columns.Add("F14", typeof(decimal));
                            dtPacklist.Columns.Add("F15", typeof(decimal));

                            invoiceno = string.Empty;
                            char[] filenamechars = filename.ToCharArray();
                            int k;
                            for (int i = 0; i < filenamechars.Length; i++)
                            {
                                if (Int32.TryParse(filenamechars[i].ToString(), out k))
                                {
                                    invoiceno = invoiceno + filenamechars[i].ToString();
                                }
                            }
                            totalrolls = Convert.ToInt32(contents[0]);
                            partyname = contents[9].Trim();
                            dispatchdate = contents[2].Trim().Replace('.', '/');
                            dispatchno = contents[3].Trim();
                            int rowcounter = 17;
                            while (totalrolls > 0)
                            {
                                string[] rowcontent = contents[rowcounter].Split();
                                object[] row = new object[] { rowcontent[0], rowcontent[1], rowcontent[2], rowcontent[4].ToLower().Contains("fresh") ? rowcontent[3] : "", rowcontent[4].ToLower().Contains("fresh") ? rowcontent[4] : rowcontent[3], rowcontent[4].ToLower().Contains("fresh") ? rowcontent[5] : rowcontent[4], "", "", "", "", "", "", rowcontent[rowcontent.Length - 3], rowcontent[rowcontent.Length - 2], rowcontent[rowcontent.Length - 1] };
                                try
                                {
                                    dtPacklist.Rows.Add(row);
                                }
                                catch (Exception ex)
                                {
                                    WriteToImportLog(filename, ex.Message, "");
                                    emailbody.AppendLine(filename + "-" + ex.Message);
                                    errorflag = 1;
                                    rowcounter++;
                                    continue;
                                }
                                if (rowcounter == 56)
                                    rowcounter = 73;
                                if (rowcounter == 113)
                                    rowcounter = 130;
                                rowcounter++;
                                totalrolls--;
                            }
                        }
                        catch (Exception ex)
                        {
                            if (filename.ToLower().Contains("inv"))
                                File.Move(files[j],
                                    Properties.Settings.Default.BhaskarProcessedLocation +
                                    Path.GetFileName(files[j]));
                            else
                            {
                                emailbody.AppendLine(filename + "-" + ex.Message);
                                WriteToImportLog(files[j], ex.Message, "");
                                errorflag = 1;
                            }
                            //unprocessedcount++;                            
                            continue;
                        }
                    }

                    #region process excel
                    else if (files[j].ToLower().EndsWith("xls"))
                    {
                        //excelConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + files[j] + ";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\"";
                        //OleDbCommand excelCommand = new OleDbCommand();
                        //OleDbDataAdapter excelDataAdapter = new OleDbDataAdapter();
                        //OleDbConnection excelConn = new OleDbConnection(excelConnStr);

                        //excelConn.Open();

                        //excelCommand = new OleDbCommand("SELECT * FROM [Sheet1$]", excelConn);

                        //excelDataAdapter.SelectCommand = excelCommand;

                        //excelDataAdapter.Fill(dtPacklist);

                        //excelConn.Close();

                        //partyname = dtPacklist.Rows[6]["F2"].ToString().Trim();
                        //dispatchdate = dtPacklist.Rows[4]["F4"].ToString().Trim().Replace('.', '/');
                        //dispatchno = dtPacklist.Rows[5]["F4"].ToString().Trim();
                        //invoiceno = filename.Substring(filename.Length - 9, 5);
                        //listcounter = 11;
                    }
                    #endregion

                    recordno = 1;
                    DateTime dispatchdt = new DateTime();
                    IFormatProvider culture = new System.Globalization.CultureInfo("fr-FR", true);
                    dispatchdt = DateTime.Parse(dispatchdate, culture);
                    while (listcounter < dtPacklist.Rows.Count)
                    {
                        try
                        {
                            if (dtPacklist.Rows[listcounter]["F1"] != System.DBNull.Value)
                            {
                                sno = dtPacklist.Rows[listcounter]["F1"].ToString().Trim();
                            }
                            else
                            {
                                break;
                            }
                            if (dtPacklist.Rows[listcounter]["F2"] != System.DBNull.Value)
                            {
                                sortno = dtPacklist.Rows[listcounter]["F2"].ToString().Trim();
                            }
                            else
                            {
                                WriteToImportLog(files[j], recordno.ToString(), "sortno");
                                emailbody.AppendLine("Sort No " + Path.GetFileName(files[j]) + recordno.ToString());
                                errorflag = 1;
                                listcounter++;
                                recordno++;
                                continue;
                            }

                            if (dtPacklist.Rows[listcounter]["F3"] != System.DBNull.Value)
                            {
                                rollno = dtPacklist.Rows[listcounter]["F3"].ToString().Trim();
                            }
                            else
                            {
                                WriteToImportLog(files[j], recordno.ToString(), "rollno");
                                emailbody.AppendLine("rollno " + Path.GetFileName(files[j]) + recordno.ToString());
                                errorflag = 1;
                                listcounter++;
                                recordno++;
                                continue;
                            }

                            if (dtPacklist.Rows[listcounter]["F4"] != System.DBNull.Value)
                            {
                                shade = dtPacklist.Rows[listcounter]["F4"].ToString().Trim();
                            }

                            if (dtPacklist.Rows[listcounter]["F5"] != System.DBNull.Value)
                            {
                                grade = dtPacklist.Rows[listcounter]["F5"].ToString().Trim();
                            }
                            else
                            {
                                WriteToImportLog(files[j], recordno.ToString(), "grade");
                                emailbody.AppendLine("grade " + Path.GetFileName(files[j]) + recordno.ToString());
                                errorflag = 1;
                                listcounter++;
                                recordno++;
                                continue;
                            }

                            if (dtPacklist.Rows[listcounter]["F6"] != System.DBNull.Value)
                            {
                                pieces = Convert.ToInt32(dtPacklist.Rows[listcounter]["F6"].ToString().Trim());
                            }
                            else
                            {
                                WriteToImportLog(files[j], recordno.ToString(), "pieces");
                                emailbody.AppendLine("pieces " + Path.GetFileName(files[j]) + recordno.ToString());
                                errorflag = 1;
                                listcounter++;
                                recordno++;
                                continue;
                            }

                            if (dtPacklist.Rows[listcounter]["F13"] != System.DBNull.Value)
                            {
                                mtrs = Convert.ToDecimal(dtPacklist.Rows[listcounter]["F13"].ToString().Trim());
                            }
                            else
                            {
                                WriteToImportLog(files[j], recordno.ToString(), "mtrs");
                                emailbody.AppendLine("mtrs " + Path.GetFileName(files[j]) + recordno.ToString());
                                errorflag = 1;
                                listcounter++;
                                recordno++;
                                continue;
                            }

                            if (dtPacklist.Rows[listcounter]["F14"] != System.DBNull.Value)
                            {
                                netweight = Convert.ToDecimal(dtPacklist.Rows[listcounter]["F14"].ToString().Trim());
                            }
                            else
                            {
                                WriteToImportLog(files[j], recordno.ToString(), "netweight");
                                emailbody.AppendLine("netweight " + Path.GetFileName(files[j]) + recordno.ToString());
                                errorflag = 1;
                                listcounter++;
                                recordno++;
                                continue;
                            }

                            if (dtPacklist.Rows[listcounter]["F15"] != System.DBNull.Value)
                            {
                                grossweight = Convert.ToDecimal(dtPacklist.Rows[listcounter]["F15"].ToString().Trim());
                            }
                            else
                            {
                                WriteToImportLog(files[j], recordno.ToString(), "grossweight");
                                emailbody.AppendLine("grossweight " + files[j] + recordno.ToString());
                                errorflag = 1;
                                listcounter++;
                                recordno++;
                                continue;
                            }

                            Entity stock = new Entity("new_stock");
                            stock["new_supplier"] = "Bhaskar";
                            stock["new_sortnumber"] = sortno;
                            stock["new_rollnumber"] = rollno;
                            stock["new_meters"] = mtrs;
                            stock["new_grade"] = grade;
                            _service.Create(stock);


                            //cmd.CommandText = @"INSERT INTO  BhaskarPackingList VALUES(" +
                            //                    sno +
                            //                    ",'" +
                            //                    sortno +
                            //                    "','" +
                            //                    rollno +
                            //                    "','" +
                            //                    shade +
                            //                    "','" +
                            //                    grade +
                            //                    "'," +
                            //                    pieces +
                            //                    "," +
                            //                    mtrs +
                            //                    "," +
                            //                    netweight +
                            //                    "," +
                            //                    grossweight +
                            //                    ",'" +
                            //                    partyname +
                            //                    "','" +
                            //                    dispatchdt.ToShortDateString() +
                            //                    "','" +
                            //                    dispatchno +
                            //                    "','','" +
                            //                    invoiceno +
                            //                    "','null',0.0,'01/01/2000','','','" +
                            //                    DateTime.Now.ToShortDateString() +
                            //                    "')";
                            //cmd.ExecuteNonQuery();

                            sno = string.Empty;
                            sortno = string.Empty;
                            rollno = string.Empty;
                            shade = string.Empty;
                            grade = string.Empty;
                            mtrs = decimal.Zero;
                            netweight = decimal.Zero;
                            grossweight = decimal.Zero;
                            pieces = 1;
                            processedcount++;
                            listcounter++;
                            recordno++;
                        }
                        catch (Exception ex)
                        {
                            WriteToImportLog(files[j], recordno.ToString(), ex.Message);
                            emailbody.AppendLine(Path.GetFileName(files[j]) + "-" + recordno.ToString() + ex.Message);
                            errorflag = 1;
                            unprocessedcount++;
                            listcounter++;
                            recordno++;
                            continue;
                        }
                    }
                    if (0 == 0)
                    {
                        File.Move(files[j], Properties.Settings.Default.BhaskarProcessedLocation + Path.GetFileName(files[j]));
                        //cmd.CommandText = "INSERT INTO ProcessedFiles VALUES('" + Properties.Settings.Default.BhaskarProcessedLocation + filename + "','" + DateTime.Now.ToShortDateString() + "')";
                        //cmd.ExecuteNonQuery();
                        filesprocessed++;
                        emailbody.AppendLine("Processed File -" + Path.GetFileName(files[j]));
                        invoiceno = string.Empty;
                    }

                    Console.WriteLine("Processed file: " + filename);
                }
                catch (Exception ex)
                {
                    WriteToImportLog(files[j], "PDF/Excel Processing Error", ex.Message);
                    //unprocessedcount++;
                    continue;
                }
            }
            //Conn.Close();
        }

        static private void WriteToDownlaodLog(string file, string error)
        {
            using (StreamWriter sw = new StreamWriter(Properties.Settings.Default.ErrorLog + file + ".txt", true))
            {
                sw.WriteLine();
                sw.WriteLine(error);
            }
        }

        static private void WriteToImportLog(string file, string row, string error)
        {
            using (StreamWriter sw = new StreamWriter(Properties.Settings.Default.ErrorLog + Path.GetFileName(file) + ".txt", true))
            {
                sw.WriteLine();
                sw.WriteLine(row + " - " + error);
            }
        }

        static private void SendMail(string subject, string body)
        {
            StringBuilder sb = new StringBuilder();
            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            message.From = new System.Net.Mail.MailAddress(Properties.Settings.Default.EmailFrom);
            message.To.Add(Properties.Settings.Default.EmailTo);
            message.Subject = subject;
            message.Body = body;
            System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new System.Net.NetworkCredential(Properties.Settings.Default.UserName, Properties.Settings.Default.Password)
            };
            smtp.Send(message);
        }

        static private string[] PDFToTextITextSharp(string file, string[] split, StringSplitOptions option)
        {
            string pdftext = string.Empty;
            PdfReader doc = new PdfReader(file);
            for (int i = 1; i <= doc.NumberOfPages; i++)
            {
                pdftext = pdftext + PdfTextExtractor.GetTextFromPage(doc, i);
            }
            doc.Close();
            return pdftext.Split(split, option);
        }

        static private string[] PDFToTextPDFBox(string file, string[] split, StringSplitOptions option)
        {
            string pdftext = string.Empty;
            PDDocument doc = PDDocument.load(file);
            PDFTextStripper stripper = new PDFTextStripper();
            pdftext = stripper.getText(doc);
            doc.close();
            return pdftext.Split(split, option);
        }

        private static IOrganizationService CreateService()
        {
            ClientCredentials Credentials = new ClientCredentials();
            ClientCredentials devivceCredentials = new ClientCredentials();

            Credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;

            //This URL needs to be updated to match the servername and Organization for the environment.

            //The following URLs should be used to access the Organization service(SOAP endpoint):
            //https://{Organization Name}.api.crm.dynamics.com/XrmServices/2011/Organization.svc (North America)
            //https://{Organization Name}.api.crm4.dynamics.com/XrmServices/2011/Organization.svc (EMEA)
            //https://{Organization Name}.api.crm5.dynamics.com/XrmServices/2011/Organization.svc (APAC)

            Uri OrganizationUri = new Uri("https://unizap.api.crm5.dynamics.com/XRMServices/2011/Organization.svc");  //Here I am using APAC.

            Uri HomeRealmUri = null;

            //To get device id and password.
            //Online: For online version, we need to call this method to get device id.
            devivceCredentials = DeviceIdManager.LoadDeviceCredentials();

            using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, devivceCredentials))
            {
                serviceProxy.ClientCredentials.UserName.UserName = "gadodia@unizap.com";  // Your Online username.Eg:username@yourcompany.onmicrosoft.com";
                serviceProxy.ClientCredentials.UserName.Password = "pass@2013";  //Your Online password
                serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());
                _service = (IOrganizationService)serviceProxy;
            }
            return _service;
        }
    }
}