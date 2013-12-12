using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Data.OleDb;
using Lesnikowski.Client;
using Lesnikowski.Client.IMAP;
using Lesnikowski.Mail;
using Lesnikowski.Mail.Fluent;
using Lesnikowski.Mail.Headers;
using Lesnikowski.Mail.Headers.Constants;
using System.Net.Mail;
using System.Net;
using org.pdfbox.pdmodel;
using org.pdfbox.util;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace DownloadMailAttachments
{
    class Program
    {
        static string filename = string.Empty;
        static int attachmentcount = 0;
        static StringBuilder packlistfiles = new StringBuilder();        
        static string connectionstring = Properties.Settings.Default.DBConnectionString;
        static string excelConnStr;
        static string sender;
        static int processedcount = 0, unprocessedcount = 0, filesprocessed = 0;
        static StringBuilder emailbody = new StringBuilder();
        static string emaildate = string.Empty;

        static void Main()
        {
            Console.WriteLine("Downloading Packlists...Please Wait...Dont close this window");
            packlistfiles.AppendLine();            
            LESNIKOWSKIMethod();            
            Console.WriteLine("Download Complete.");
            //ProcessBhaskarPacklists();
            //ProcessKGPacklists();
            //try
            //{
            //    SendMail("Import Log", "Files Processed: " + filesprocessed.ToString() + "\n\rRecords: Processed " + processedcount + " / Unprocessed " + unprocessedcount + " " + emailbody.ToString());
            //}
            //catch (Exception ex)
            //{
            //    WriteToImportLog(filename, "-1", "Mail not sent");
            //}
            //Console.Write("Process Complete...Press any key to continue");
            //System.Threading.Thread.Sleep(2000);
        }        

        static private void LESNIKOWSKIMethod()
        {
            using (Imap imap = new Imap())
            {
                try
                {
                    imap.Connect("imap.gmail.com", 993, true);
                    imap.UseBestLogin(Properties.Settings.Default.UserName, Properties.Settings.Default.Password);
                    imap.SelectInbox();
                    List<long> uidList = imap.SearchFlag(Flag.Unseen);
                    foreach (long uid in uidList)
                    {
                        try
                        {
                            IMail email = new MailBuilder()
                                .CreateFromEml(imap.GetMessageByUID(uid));                           
                            attachmentcount = 0;
                            emaildate = email.Date.ToString();
                            foreach (MailBox m in email.From)
                            {
                                if (m.Address.ToLower().Contains("kg"))
                                {
                                    sender = "kg";
                                    break;
                                }
                                else if (m.Address.ToLower().Contains("bhaskar"))
                                {
                                    sender = "bhaskar";
                                    break;
                                }
                            }
                            foreach (MimeData attachment in email.Attachments)
                            {
                                filename = attachment.SafeFileName;
                                if (sender == "bhaskar")
                                    attachment.Save(Properties.Settings.Default.BhaskarPacklistSavePath + attachment.SafeFileName);
                                else if (sender == "kg")
                                    attachment.Save(Properties.Settings.Default.KGPacklistSavePath + attachment.SafeFileName);
                                attachmentcount++;
                                packlistfiles.AppendLine(attachment.SafeFileName); 
                            }
                            Console.WriteLine("Packlists of " + email.Date + " saved");
                            SendMail("Download Log", attachmentcount.ToString() +
                                " Packlist files dt. " +
                                emaildate +
                                " downloaded." +
                                packlistfiles.ToString() +
                                "\r\n");
                        }
                        catch (Exception ex)
                        {
                            WriteToDownlaodLog(filename, ex.Message);
                            continue;
                        }
                    }
                    imap.Close();
                }
                catch (Exception ex)
                {
                    WriteToDownlaodLog(filename, ex.Message);
                }
            }
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
            OleDbConnection Conn = new OleDbConnection(connectionstring);
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = Conn;
            cmd.CommandType = CommandType.Text;
            Conn.Open();

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
                        excelConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + files[j] + ";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\"";
                        OleDbCommand excelCommand = new OleDbCommand();
                        OleDbDataAdapter excelDataAdapter = new OleDbDataAdapter();
                        OleDbConnection excelConn = new OleDbConnection(excelConnStr);

                        excelConn.Open();

                        excelCommand = new OleDbCommand("SELECT * FROM [Sheet1$]", excelConn);

                        excelDataAdapter.SelectCommand = excelCommand;

                        excelDataAdapter.Fill(dtPacklist);

                        excelConn.Close();

                        partyname = dtPacklist.Rows[6]["F2"].ToString().Trim();
                        dispatchdate = dtPacklist.Rows[4]["F4"].ToString().Trim().Replace('.', '/');
                        dispatchno = dtPacklist.Rows[5]["F4"].ToString().Trim();
                        invoiceno = filename.Substring(filename.Length - 9, 5);
                        listcounter = 11;
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

                            cmd.CommandText =  @"INSERT INTO  BhaskarPackingList VALUES(" + 
                                                sno + 
                                                ",'" + 
                                                sortno + 
                                                "','" + 
                                                rollno + 
                                                "','" + 
                                                shade + 
                                                "','" + 
                                                grade + 
                                                "'," + 
                                                pieces + 
                                                "," + 
                                                mtrs + 
                                                "," +
                                                netweight + 
                                                "," + 
                                                grossweight + 
                                                ",'" + 
                                                partyname + 
                                                "','" + 
                                                dispatchdt.ToShortDateString() + 
                                                "','" + 
                                                dispatchno + 
                                                "','','" + 
                                                invoiceno + 
                                                "','null',0.0,'01/01/2000','','','" + 
                                                DateTime.Now.ToShortDateString() + 
                                                "')";
                            cmd.ExecuteNonQuery();

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
                    if (0 == 0 )
                    {
                        File.Move(files[j], Properties.Settings.Default.BhaskarProcessedLocation + Path.GetFileName(files[j]));
                        cmd.CommandText = "INSERT INTO ProcessedFiles VALUES('" + Properties.Settings.Default.BhaskarProcessedLocation + filename + "','" + DateTime.Now.ToShortDateString() + "')";
                        cmd.ExecuteNonQuery();
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
            Conn.Close();            
        }

        private static void ProcessKGPacklists()
        {
            //Table Columns
            string
                docno = string.Empty,
                direfno = string.Empty,
                sortno = string.Empty,
                rollno = string.Empty,
                shade = string.Empty,
                grade = string.Empty,
                pieceno = string.Empty,
                inwardno = string.Empty,
                soldto = string.Empty,
                dispatchdate = string.Empty,
                saledate = string.Empty,
                challanreference = string.Empty,
                datemodified = string.Empty;
            int
                pieces = 1;
            decimal
                meters = decimal.Zero,
                netweight = decimal.Zero,
                grossweight = decimal.Zero;

            //Other variables            
            string[] split = new string[] { "\n" };
            int recordno, errorflag;
            emailbody.AppendLine();

            DataTable dtPacklist = new DataTable();
            dtPacklist.TableName = "Packlist";

            OleDbConnection conn = new OleDbConnection(connectionstring);
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandType = CommandType.Text;
            conn.Open();

            string[] files = Directory.GetFiles(Properties.Settings.Default.KGPacklistSavePath);
            for (int j = 0; j < files.Length; j++)
            {
                try
                {
                    if (files[j].ToLower().EndsWith("pdf"))
                    {
                        int totalrolls = 0;
                        filename = Path.GetFileName(files[j]);
                        Console.WriteLine("Processing File: " + filename);
                        int listcounter = 0;
                        errorflag = 0;
                        dtPacklist.Reset();
                        try
                        {
                            string[] contents = PDFToTextITextSharp(files[j], split, StringSplitOptions.None);

                            dtPacklist.Columns.Add("F1", typeof(string));
                            dtPacklist.Columns.Add("F2", typeof(string));
                            dtPacklist.Columns.Add("F3", typeof(string));
                            dtPacklist.Columns.Add("F4", typeof(string));
                            dtPacklist.Columns.Add("F5", typeof(string));
                            dtPacklist.Columns.Add("F6", typeof(string));
                            dtPacklist.Columns.Add("F7", typeof(string));
                            dtPacklist.Columns.Add("F8", typeof(decimal));
                            dtPacklist.Columns.Add("F9", typeof(decimal));
                            dtPacklist.Columns.Add("F10", typeof(decimal));
                            dtPacklist.Columns.Add("F11", typeof(decimal));

                            dispatchdate = contents[5].Trim().Substring(7, 10);                            
                            string footer = contents[contents.Length - 4];
                            int.TryParse(footer.Substring(footer.IndexOf('(') + 1, footer.IndexOf(')') - footer.IndexOf('(') - 1), out totalrolls);
                            int rowcounter = contents.Length - 7;
                            int[] indicesoftotal = new int[100];
                            int index = 0;
                            for (int i = contents.Length - 5; i > 0; i--)
                            {
                                if (contents[i].Equals("Total"))
                                    indicesoftotal[index++] = i;
                            }
                            indicesoftotal[index] = 11;
                            for (int i = index; i > 0; i--)
                            {
                                for (int k = indicesoftotal[i] + 1; k < indicesoftotal[i - 1]; k++)
                                {
                                    if (contents[k].ToLower().Contains("prepared"))
                                    {
                                        k += 12;
                                        continue;
                                    }
                                    else
                                    {
                                        string[] contentrow = new string[contents[k].Split(' ').Length];
                                        contentrow = contents[k].Split(' ');
                                        if (contentrow.Length == 2)
                                        {
                                            if (contentrow[0].Length > 1)
                                                rollno = contentrow[0] + contentrow[1];
                                        }
                                        else if (contentrow.Length == 16)//first roll of the sub-lot
                                        {
                                            docno = contentrow[1];
                                            direfno = contentrow[3];
                                            sortno = contentrow[4] + " " + contentrow[5];
                                            grade = contentrow[6];
                                            shade = contentrow[7];
                                            pieceno = contentrow[8];
                                            pieces = int.Parse(contentrow[9]);
                                            meters = decimal.Parse(contentrow[12]);
                                            netweight = decimal.Parse(contentrow[14]);
                                            grossweight = decimal.Parse(contentrow[15]);
                                        }
                                        else if (contentrow.Length == 15)
                                        {
                                            docno = contentrow[1];
                                            direfno = contentrow[3];
                                            sortno = contentrow[4] + " " + contentrow[5];
                                            grade = contentrow[6];
                                            if (contentrow[10].Contains('(') && contentrow[10].Contains(')'))
                                            {
                                                shade = contentrow[7];
                                                pieceno = contentrow[8];
                                                pieces = int.Parse(contentrow[9]);
                                                meters = decimal.Parse(contentrow[11]);
                                                netweight = decimal.Parse(contentrow[13]);
                                                grossweight = decimal.Parse(contentrow[14]);
                                            }
                                            else//first 
                                            {
                                                shade = "";
                                                pieceno = contentrow[7];
                                                pieces = int.Parse(contentrow[8]);
                                                meters = decimal.Parse(contentrow[11]);
                                                netweight = decimal.Parse(contentrow[13]);
                                                grossweight = decimal.Parse(contentrow[14]);
                                            }
                                        }
                                        else if (contentrow.Length == 17)
                                        {
                                            docno = contentrow[1];
                                            direfno = contentrow[3];
                                            sortno = contentrow[4] + " " + contentrow[5];
                                            grade = contentrow[6];
                                            rollno = contentrow[7] + contentrow[8];
                                            pieceno = contentrow[9];
                                            pieces = int.Parse(contentrow[10]);
                                            meters = decimal.Parse(contentrow[13]);
                                            netweight = decimal.Parse(contentrow[15]);
                                            grossweight = decimal.Parse(contentrow[16]);
                                        }
                                        else if (contentrow.Length == 18)
                                        {
                                            docno = contentrow[1];
                                            direfno = contentrow[3];
                                            sortno = contentrow[4] + " " + contentrow[5];
                                            grade = contentrow[6];
                                            shade = contentrow[7];
                                            rollno = contentrow[8] + contentrow[9];
                                            pieceno = contentrow[10];
                                            pieces = int.Parse(contentrow[11]);
                                            meters = decimal.Parse(contentrow[14]);
                                            netweight = decimal.Parse(contentrow[16]);
                                            grossweight = decimal.Parse(contentrow[17]);
                                        }
                                        else if (contentrow.Length == 10)
                                        {
                                            rollno = contentrow[0] + contentrow[1];
                                            pieceno = contentrow[2];
                                            pieces = int.Parse(contentrow[3]);
                                            meters = decimal.Parse(contentrow[6]);
                                            netweight = decimal.Parse(contentrow[8]);
                                            grossweight = decimal.Parse(contentrow[9]);
                                        }
                                        else if (contentrow.Length == 11)
                                        {
                                            shade = contentrow[0];
                                            rollno = contentrow[1] + contentrow[2];
                                            pieceno = contentrow[3];
                                            pieces = int.Parse(contentrow[4]);
                                            meters = decimal.Parse(contentrow[7]);
                                            netweight = decimal.Parse(contentrow[9]);
                                            grossweight = decimal.Parse(contentrow[10]);
                                        }
                                        else if (contentrow.Length == 9)
                                        {
                                            shade = contentrow[0];
                                            pieceno = contentrow[1];
                                            pieces = int.Parse(contentrow[2]);
                                            meters = decimal.Parse(contentrow[5]);
                                            netweight = decimal.Parse(contentrow[7]);
                                            grossweight = decimal.Parse(contentrow[8]);
                                        }
                                        else if (contentrow.Length == 8)
                                        {
                                            shade = "";
                                            pieceno = contentrow[0];
                                            pieces = int.Parse(contentrow[1]);
                                            meters = decimal.Parse(contentrow[4]);
                                            netweight = decimal.Parse(contentrow[6]);
                                            grossweight = decimal.Parse(contentrow[7]);
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                        if (contentrow.Length != 2)
                                        {
                                            object[] row = new object[]
                                            {
                                                docno,
                                                direfno,
                                                sortno,
                                                rollno,
                                                shade,
                                                grade,
                                                pieceno,
                                                pieces,
                                                meters,
                                                netweight,
                                                grossweight
                                            };
                                            dtPacklist.Rows.Add(row);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteToImportLog(files[j], ex.Message, "");
                            errorflag = 1;
                            //unprocessedcount++;
                            emailbody.AppendLine(filename + "-" + ex.Message);
                            continue;
                        }
                        recordno = 1;
                        DateTime dispatchdt = new DateTime();
                        IFormatProvider culture = new System.Globalization.CultureInfo("fr-FR", true);
                        dispatchdt = DateTime.Parse(dispatchdate, culture);
                        int rowsinserted = 0;
                        while (listcounter < dtPacklist.Rows.Count)
                        {
                            try
                            {
                                if (dtPacklist.Rows[listcounter]["F1"] != System.DBNull.Value)
                                {
                                    docno = dtPacklist.Rows[listcounter]["F1"].ToString().Trim();
                                }
                                else
                                {
                                    WriteToImportLog(files[j], recordno.ToString(), "doc no");
                                    emailbody.AppendLine("doc no " + filename + recordno.ToString());
                                    errorflag = 1;
                                    listcounter++;
                                    recordno++;
                                    continue;
                                }
                                if (dtPacklist.Rows[listcounter]["F2"] != System.DBNull.Value)
                                {
                                    direfno = dtPacklist.Rows[listcounter]["F2"].ToString().Trim();
                                }
                                else
                                {
                                    WriteToImportLog(files[j], recordno.ToString(), "direfno");
                                    emailbody.AppendLine("DI REF NO " + filename + recordno.ToString());
                                    errorflag = 1;
                                    listcounter++;
                                    recordno++;
                                    continue;
                                }

                                if (dtPacklist.Rows[listcounter]["F3"] != System.DBNull.Value)
                                {
                                    sortno = dtPacklist.Rows[listcounter]["F3"].ToString().Trim();
                                }
                                else
                                {
                                    WriteToImportLog(files[j], recordno.ToString(), "sortnno");
                                    emailbody.AppendLine("sortno " + filename + recordno.ToString());
                                    errorflag = 1;
                                    listcounter++;
                                    recordno++;
                                    continue;
                                }

                                if (dtPacklist.Rows[listcounter]["F4"] != System.DBNull.Value)
                                {
                                    rollno = dtPacklist.Rows[listcounter]["F4"].ToString().Trim();
                                }
                                else
                                {
                                    WriteToImportLog(files[j], recordno.ToString(), "rollno");
                                    emailbody.AppendLine("rollno " + filename + recordno.ToString());
                                    errorflag = 1;
                                    listcounter++;
                                    recordno++;
                                    continue;
                                }

                                if (dtPacklist.Rows[listcounter]["F5"] != System.DBNull.Value)
                                {
                                    shade = dtPacklist.Rows[listcounter]["F5"].ToString().Trim();
                                }


                                if (dtPacklist.Rows[listcounter]["F6"] != System.DBNull.Value)
                                {
                                    grade = dtPacklist.Rows[listcounter]["F6"].ToString().Trim();
                                }
                                else
                                {
                                    WriteToImportLog(files[j], recordno.ToString(), "grade");
                                    emailbody.AppendLine("grade " + filename + recordno.ToString());
                                    errorflag = 1;
                                    listcounter++;
                                    recordno++;
                                    continue;
                                }

                                if (dtPacklist.Rows[listcounter]["F7"] != System.DBNull.Value)
                                {
                                    pieceno = dtPacklist.Rows[listcounter]["F7"].ToString().Trim();
                                }
                                else
                                {
                                    WriteToImportLog(files[j], recordno.ToString(), "pieceno");
                                    emailbody.AppendLine("pieceno " + filename + recordno.ToString());
                                    errorflag = 1;
                                    listcounter++;
                                    recordno++;
                                    continue;
                                }

                                if (dtPacklist.Rows[listcounter]["F8"] != System.DBNull.Value)
                                {
                                    pieces = Convert.ToInt32(dtPacklist.Rows[listcounter]["F8"].ToString().Trim());
                                }
                                else
                                {
                                    WriteToImportLog(files[j], recordno.ToString(), "pieces");
                                    emailbody.AppendLine("pieces " + filename + recordno.ToString());
                                    errorflag = 1;
                                    listcounter++;
                                    recordno++;
                                    continue;
                                }

                                if (dtPacklist.Rows[listcounter]["F9"] != System.DBNull.Value)
                                {
                                    meters = Convert.ToDecimal(dtPacklist.Rows[listcounter]["F9"].ToString().Trim());
                                }
                                else
                                {
                                    WriteToImportLog(files[j], recordno.ToString(), "meters");
                                    emailbody.AppendLine("meters " + filename + recordno.ToString());
                                    errorflag = 1;
                                    listcounter++;
                                    recordno++;
                                    continue;
                                }
                                if (dtPacklist.Rows[listcounter]["F10"] != System.DBNull.Value)
                                {
                                    netweight = Convert.ToDecimal(dtPacklist.Rows[listcounter]["F10"].ToString().Trim());
                                }
                                else
                                {
                                    WriteToImportLog(files[j], recordno.ToString(), "netweight");
                                    emailbody.AppendLine("netweight " + filename + recordno.ToString());
                                    errorflag = 1;
                                    listcounter++;
                                    recordno++;
                                    continue;
                                }
                                if (dtPacklist.Rows[listcounter]["F11"] != System.DBNull.Value)
                                {
                                    grossweight = Convert.ToDecimal(dtPacklist.Rows[listcounter]["F11"].ToString().Trim());
                                }
                                else
                                {
                                    WriteToImportLog(files[j], recordno.ToString(), "grossweight");
                                    emailbody.AppendLine("grossweight " + filename + recordno.ToString());
                                    errorflag = 1;
                                    listcounter++;
                                    recordno++;
                                    continue;
                                }

                                cmd.CommandText = @"INSERT INTO KGPackingList " +
                                                   "VALUES ('" +
                                                   docno +
                                                   "','" +
                                                   direfno +
                                                   "','" +
                                                   sortno +
                                                   "','" +
                                                   rollno +
                                                   "','" +
                                                   shade +
                                                   "','" +
                                                   grade +
                                                   "','" +
                                                   pieceno +
                                                   "'," +
                                                   pieces +
                                                   "," +
                                                   meters +
                                                   "," +
                                                   netweight +
                                                   "," +
                                                   grossweight +
                                                   ",'" +
                                                   dispatchdt.ToShortDateString() +
                                                   "','" +
                                                   "','null',0.0,'1/1/2000','','" +
                                                   DateTime.Now.ToShortDateString() +
                                                   "')";

                                cmd.ExecuteNonQuery();

                                docno = string.Empty;
                                direfno = string.Empty;
                                sortno = string.Empty;
                                rollno = string.Empty;
                                shade = string.Empty;
                                grade = string.Empty;
                                pieces = 1;
                                pieceno = string.Empty;
                                meters = decimal.Zero;
                                netweight = decimal.Zero;
                                grossweight = decimal.Zero;
                                pieces = 1;
                                processedcount++;
                                listcounter++;
                                recordno++;
                                rowsinserted++;
                            }
                            catch (Exception ex)
                            {
                                WriteToImportLog(files[j], recordno.ToString(), ex.Message);
                                emailbody.AppendLine(filename + "-" + recordno.ToString() + ex.Message);
                                errorflag = 1;
                                unprocessedcount++;
                                listcounter++;
                                recordno++;
                                continue;
                            }
                        }
                        if (errorflag == 0 || rowsinserted == totalrolls)
                        {
                            File.Move(files[j], Properties.Settings.Default.KGProcessedLocation + filename);
                            cmd.CommandText = @"INSERT INTO ProcessedFiles VALUES('" +
                                                Properties.Settings.Default.KGProcessedLocation +
                                                filename +
                                                "','" +
                                                DateTime.Now.ToShortDateString() + "')";
                            cmd.ExecuteNonQuery();
                            filesprocessed++;
                            emailbody.AppendLine("Processed File -" + filename);
                        }
                        Console.WriteLine("Processed file: " + filename);
                    }
                }
                catch (Exception ex)
                {
                    WriteToImportLog(files[j], "PDF Processing Error", ex.Message);
                    //unprocessedcount++;
                    continue;
                }
            }            
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

        /*
      static private void IMAPXMethod()
      {            
          try
          {
              ImapX.ImapClient client = new ImapX.ImapClient("imap.gmail.com", 993, true);
                
              bool result = false;

              result = client.Connection();

              result = client.LogIn(Properties.Settings.Default.UserName, Properties.Settings.Default.Password);

              ImapX.MessageCollection mc = client.Folders["INBOX"].Search("UNSEEN", true);

              foreach (ImapX.Message m in mc)
              {
                  packlistcount = 0;
                  List<ImapX.Attachment> attachments = m.Attachments;

                  foreach (ImapX.Attachment attachment in attachments)
                  {

                      if (!attachment.FileName.Contains("inv"))
                      {
                          filename = attachment.FileName;
                          attachment.SaveFile(Properties.Settings.Default.SavePath);
                          packlistcount++;
                      }
                  }                    
                  m.SetFlag(ImapX.ImapFlags.SEEN);
                  Console.WriteLine("Packlists of " + m.Date.ToString() + " saved");
                  SendMail(m.Date.ToString());
              }
              client.LogOut();
              client.Disconnect();
          }
          catch (Exception ex)
          {
              WriteToDownlaodLog(filename, ex.Message);
              SendMail(filename + "-" + ex.Message);
          }
      }*/
    }
}
