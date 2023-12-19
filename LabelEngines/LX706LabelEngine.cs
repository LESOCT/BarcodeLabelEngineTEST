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
    public class LX706LabelEngine : IDisposable
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
            LX706LabelReareangeSettings LX706RearrangeSettings = new LX706LabelReareangeSettings();
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                    string Customer_Item_Number_Barcode = "";
                    try
                    {
                        Customer_Item_Number_Barcode = lofLines[LX706RearrangeSettings.Customer_Item_Number_Barcode.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.Customer_Item_Number_Barcode.RearrangeColumnStart, LX706RearrangeSettings.Customer_Item_Number_Barcode.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Customer_Item_Number_Barcode = lofLines[LX706RearrangeSettings.Customer_Item_Number_Barcode.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.Customer_Item_Number_Barcode.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }
                    string Item_Number = "";
                    try
                    {
                        Item_Number = lofLines[LX706RearrangeSettings.Item_Number.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.Item_Number.RearrangeColumnStart, LX706RearrangeSettings.Item_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number = lofLines[LX706RearrangeSettings.Item_Number.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.Item_Number.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }





                    string Edited_Toyta_Cust_item_no = "";
                    try
                    {
                        Edited_Toyta_Cust_item_no = lofLines[LX706RearrangeSettings.Edited_Toyta_Cust_item_no.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.Edited_Toyta_Cust_item_no.RearrangeColumnStart, LX706RearrangeSettings.Edited_Toyta_Cust_item_no.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Edited_Toyta_Cust_item_no = lofLines[LX706RearrangeSettings.Edited_Toyta_Cust_item_no.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.Edited_Toyta_Cust_item_no.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string IXDESC_and_IXDES2 = "";
                    try
                    {
                        IXDESC_and_IXDES2 = lofLines[LX706RearrangeSettings.IXDESC_and_IXDES2.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.IXDESC_and_IXDES2.RearrangeColumnStart, LX706RearrangeSettings.IXDESC_and_IXDES2.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            IXDESC_and_IXDES2 = lofLines[LX706RearrangeSettings.IXDESC_and_IXDES2.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.IXDESC_and_IXDES2.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }


                    string User = "";
                    try
                    {
                        User = lofLines[LX706RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.User.RearrangeColumnStart, LX706RearrangeSettings.User.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            User = lofLines[LX706RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.User.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string BatchFirst = "";
                    try
                    {
                        BatchFirst = lofLines[LX706RearrangeSettings.BatchFirst.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.BatchFirst.RearrangeColumnStart, LX706RearrangeSettings.BatchFirst.RearrangeColumnEnd).Trim();
                        try { BatchFirst = Convert.ToInt32(BatchFirst).ToString(); } catch { }
                    }
                    catch
                    {
                        try
                        {
                            BatchFirst = lofLines[LX706RearrangeSettings.BatchFirst.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.BatchFirst.RearrangeColumnStart).Trim();
                            try { BatchFirst = Convert.ToInt32(BatchFirst).ToString(); } catch { }
                        }
                        catch
                        {

                        }
                    }

                    string DateTime_Nows = "";
                    try
                    {
                        DateTime_Nows = lofLines[LX706RearrangeSettings.DateTime_Now.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.DateTime_Now.RearrangeColumnStart, LX706RearrangeSettings.DateTime_Now.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            DateTime_Nows = lofLines[LX706RearrangeSettings.DateTime_Now.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.DateTime_Now.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Printer_IP = "";
                    try
                    {
                        Printer_IP = lofLines[LX706RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.Printer_IP.RearrangeColumnStart, LX706RearrangeSettings.Printer_IP.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Printer_IP = lofLines[LX706RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.Printer_IP.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string NumberOfCopies = "";
                    try
                    {
                        NumberOfCopies = lofLines[LX706RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.NumberOfCopies.RearrangeColumnStart, LX706RearrangeSettings.NumberOfCopies.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            NumberOfCopies = lofLines[LX706RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.NumberOfCopies.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    //string BatchSecond = "";
                    //try
                    //{
                    //    BatchSecond = lofLines[LX706RearrangeSettings.BatchSecond.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.BatchSecond.RearrangeColumnStart, LX706RearrangeSettings.BatchSecond.RearrangeColumnEnd).Trim();
                    //}
                    //catch
                    //{
                    //    BatchSecond = lofLines[LX706RearrangeSettings.BatchSecond.RearrangeRowNumber - 1].Substring(LX706RearrangeSettings.BatchSecond.RearrangeColumnStart).Trim();
                    //}

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\706-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Found Template");
                            string Edited_Toyta_Cust_item_no_Cust_item_number, IXDESC__IXDES2, Users, Batch_Num2, Batch_Num1, DateTime_now;

                            Edited_Toyta_Cust_item_no_Cust_item_number = "ItemNum";
                            IXDESC__IXDES2 = "ItemDescription";
                            Users = "Users";
                            Batch_Num1 = "Bch1";
                            Batch_Num2 = "Bch2";
                            DateTime_now = "ToDate";

                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "706-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Opened New Template");
                                documents.Replace(Edited_Toyta_Cust_item_no_Cust_item_number, Edited_Toyta_Cust_item_no, false, true);
                                documents.Replace(IXDESC__IXDES2, IXDESC_and_IXDES2, false, true);
                                documents.Replace(Batch_Num1, BatchFirst, false, true);
                                documents.Replace(Batch_Num2, " ", false, true);
                                documents.Replace(DateTime_now, DateTime_Nows, false, true);
                                documents.Replace(Users, User, false, true);
                                documents.Save(wordTemplate);
                                documents.Close();
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "706-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Start Barcode Insert");

                                        PdfPage page = pdfDocument.Pages[0];
                                        PdfCode39Barcode barcode = new PdfCode39Barcode();

                                        barcode.BarHeight = 20;
                                        barcode.Text = Customer_Item_Number_Barcode;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Size = new SizeF(160, 20);
                                        barcode.Draw(page, new PointF(5, 45));

                                        PdfCode39Barcode barcode1 = new PdfCode39Barcode();

                                        barcode1.BarHeight = 20;
                                        barcode1.Text = Item_Number;
                                        barcode1.TextDisplayLocation = TextLocation.None;
                                        barcode1.Size = new SizeF(160, 20);
                                        barcode1.Draw(page, new PointF(5, 175));



                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "706-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "706-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "706-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "706-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }
    }

    public class LX706RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LX706LabelReareangeSettings
    {
        public RearrangeSetting NumberOfCopies
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 10,
                    RearrangeRowNumber = 4
                };
            }
        }

        public RearrangeSetting Printer_IP
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 44,
                    RearrangeRowNumber = 3
                };
            }
        }

        public LX706RearrangeSetting Edited_Toyta_Cust_item_no
        {
            get
            {
                return new LX706RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 45,
                    RearrangeRowNumber = 29
                };
            }
        }

        public LX706RearrangeSetting IXDESC_and_IXDES2
        {
            get
            {
                return new LX706RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 67,
                    RearrangeRowNumber = 16
                };
            }
        }



        public LX706RearrangeSetting User
        {
            get
            {
                return new LX706RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 37,
                    RearrangeRowNumber = 19
                };
            }
        }
        public LX706RearrangeSetting Item_Number
        {
            get
            {
                return new LX706RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 46,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX706RearrangeSetting Customer_Item_Number_Barcode
        {
            get
            {
                return new LX706RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 46,
                    RearrangeRowNumber = 13
                };
            }
        }

        public LX706RearrangeSetting BatchFirst
        {
            get
            {
                return new LX706RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 21,
                    RearrangeRowNumber = 21
                };
            }
        }
        //public LX706RearrangeSetting BatchSecond
        //{
        //    get
        //    {
        //        return new LX706RearrangeSetting()
        //        {
        //            RearrangeColumnStart = 35,
        //            RearrangeColumnEnd = 57,
        //            RearrangeRowNumber = 21
        //        };
        //    }
        //}

        public LX706RearrangeSetting DateTime_Now
        {
            get
            {
                return new LX706RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 5,
                    RearrangeRowNumber = 22
                };
            }
        }


    }
}

    

