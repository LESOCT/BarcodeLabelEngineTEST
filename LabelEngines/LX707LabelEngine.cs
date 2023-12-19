using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Syncfusion.Pdf.Barcode;
using System.Drawing;
using Syncfusion.Pdf;
using System.Configuration;
using Syncfusion.Pdf.Parsing;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;


namespace BarcodeLabelSoftware
{
    public class LX707LabelEngine : IDisposable
    {
        bool disposed = false;
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                LogEngine logEngine = new LogEngine(); logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Tasks", "Disposed Label Engine"); handle.Dispose();
            }

            disposed = true;
        }
        public List<FileInfo> lofFileData { get; set; }

        public async Task GenerateLabel()
        {
            await Task.Run(() => Start());
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public void Start()
        {
            LX707LabelReareangeSettings LX707RearrangeSettings = new LX707LabelReareangeSettings();
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
            while (lofFileData.Count > 0)
            {
                try {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);



                    string Customer_Item_Number = "";
                    try
                    {
                        Customer_Item_Number = lofLines[LX707RearrangeSettings.Customer_Item_Number.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Customer_Item_Number.RearrangeColumnStart, LX707RearrangeSettings.Customer_Item_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        Customer_Item_Number = lofLines[LX707RearrangeSettings.Customer_Item_Number.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Customer_Item_Number.RearrangeColumnStart).TrimStart();
                    }






                    string Batch = "";
                    try
                    {
                        Batch = lofLines[LX707RearrangeSettings.Batch.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Batch.RearrangeColumnStart, LX707RearrangeSettings.Batch.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        Batch = lofLines[LX707RearrangeSettings.Batch.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Batch.RearrangeColumnStart).Trim();
                    }



                    string IXDESC = "";
                    try
                    {
                        IXDESC = lofLines[LX707RearrangeSettings.IXDESC.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.IXDESC.RearrangeColumnStart, LX707RearrangeSettings.IXDESC.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        IXDESC = lofLines[LX707RearrangeSettings.IXDESC.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.IXDESC.RearrangeColumnStart).TrimStart();
                    }







                    string Date_Encrypted = "";
                    try
                    {
                        Date_Encrypted = lofLines[LX707RearrangeSettings.Date_Encrypted.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Date_Encrypted.RearrangeColumnStart, LX707RearrangeSettings.Date_Encrypted.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        Date_Encrypted = lofLines[LX707RearrangeSettings.Date_Encrypted.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Date_Encrypted.RearrangeColumnStart).Trim();
                    }




                    string User = "";
                    try
                    {
                        User = lofLines[LX707RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.User.RearrangeColumnStart, LX707RearrangeSettings.User.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        User = lofLines[LX707RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.User.RearrangeColumnStart).TrimStart();
                    }

                    string Printer_IP = "";
                    try
                    {
                        Printer_IP = lofLines[LX707RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Printer_IP.RearrangeColumnStart, LX707RearrangeSettings.Printer_IP.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Printer_IP = lofLines[LX707RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Printer_IP.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string NumberOfCopies = "";
                    try
                    {
                        NumberOfCopies = lofLines[LX707RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.NumberOfCopies.RearrangeColumnStart, LX707RearrangeSettings.NumberOfCopies.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            NumberOfCopies = lofLines[LX707RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.NumberOfCopies.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\707-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Found Template");
                            string txtBatch, txtIXDESC, txtToTime, txtDateEncrypted, txtCustomer_Item_Number, txtToDate, txtUser;
                            DateTime D = DateTime.Now;


                            txtCustomer_Item_Number = "CustNum";
                            txtBatch = "Btch";
                            txtIXDESC = "IXDES2";
                            txtUser = "Users";
                            txtToDate = "ToDate";
                            txtToTime = "ToTime";
                            txtDateEncrypted = "DateEncry";








                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "707-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {


                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Opened New Template");
                                documents.Replace(txtCustomer_Item_Number, Customer_Item_Number, false, true);
                                documents.Replace(txtBatch, Batch, false, true);
                                documents.Replace(txtDateEncrypted, Date_Encrypted, false, true);
                                documents.Replace(txtIXDESC, IXDESC, false, true);
                                documents.Replace(txtToDate, D.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(txtToTime, D.ToString("HH:mm:ss"), false, true);
                                documents.Replace(txtUser, User, false, true);



                                // ======== QUANTITY DOESNT WANT TO CONVERT 



                                documents.Save(wordTemplate);
                                documents.Close();
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "707-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Start Barcode Insert");

                                        PdfPage page = pdfDocument.Pages[0];

                                        PdfCode39Barcode barcode1 = new PdfCode39Barcode();
                                        barcode1.BarHeight = 25;
                                        barcode1.Text = Customer_Item_Number;
                                        barcode1.TextDisplayLocation = TextLocation.None;
                                        barcode1.Size = new SizeF(180, 25);
                                        barcode1.Draw(page, new PointF(15, 10));
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Failed to Insert Barcode - Error " + ex.ToString());
                                    }

                                    pdfDocument.Save(newPDFFileName);
                                    pdfDocument.Close(true);
                                }
                            }

                            int totalNumberOdPages = 1;
                            try
                            {
                                totalNumberOdPages = Convert.ToInt32(NumberOfCopies);
                            }
                            catch
                            {

                            }

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "707-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
                            File.Copy(newPDFFileName, outputPDFFile);
                            //PdfLoadedDocument loadedDocument = new PdfLoadedDocument(newPDFFileName);
                            //PdfDocument document = new PdfDocument();
                            //for (int i = 1; i <= totalNumberOdPages; i++)
                            //{
                            //    document.ImportPageRange(loadedDocument, 0, 0);
                            //}
                            //document.Save(outputPDFFile);
                            //document.Close(true);
                            //loadedDocument.Close(true);


                            File.Delete(wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "707-LX");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
                catch (Exception ex)
                {
                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "707-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "707-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }
    }
}

    public class LX707RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LX707LabelReareangeSettings
    {
        public LX707RearrangeSetting NumberOfCopies
        {
            get
            {
                return new LX707RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 10,
                    RearrangeRowNumber = 4
                };
            }
        }

        public LX707RearrangeSetting Printer_IP
        {
            get
            {
                return new LX707RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 44,
                    RearrangeRowNumber = 3
                };
            }
        }

        public LX707RearrangeSetting Customer_Item_Number
        {
            get
            {
                return new LX707RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 48,
                    RearrangeRowNumber = 12
                };
            }
        }

        public LX707RearrangeSetting Batch
        {
            get
            {
                return new LX707RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 40,
                    RearrangeRowNumber = 21
                };
            }
        }

        public LX707RearrangeSetting Date_Encrypted
        {
            get
            {
                return new LX707RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 34,
                    RearrangeRowNumber = 23
                };
            }
        }

        public LX707RearrangeSetting IXDESC
        {
            get
            {
                return new LX707RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 60,
                    RearrangeRowNumber = 14
                };
            }
        }







        // ======= NEEDS TO BE CHANGED IN TO A INT ===============

        public LX707RearrangeSetting ToDates
        {
            get
            {
                return new LX707RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 22
                };
            }
        }


        public LX707RearrangeSetting User
        {
            get
            {
                return new LX707RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 19
                };
            }
        }

    }

