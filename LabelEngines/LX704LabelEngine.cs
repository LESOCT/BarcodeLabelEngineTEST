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
    public class LX704LabelEngine : IDisposable
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
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
            LabelReareangeSettings rearrangeSettings = new LabelReareangeSettings();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                    string IXITEM = "";
                    try
                    {
                        IXITEM = lofLines[rearrangeSettings.IXITEM.RearrangeRowNumber - 1].Substring(rearrangeSettings.IXITEM.RearrangeColumnStart, rearrangeSettings.IXITEM.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            IXITEM = lofLines[rearrangeSettings.IXITEM.RearrangeRowNumber - 1].Substring(rearrangeSettings.IXITEM.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string IXDESC = "";
                    try
                    {
                        IXDESC = lofLines[rearrangeSettings.IXDESC.RearrangeRowNumber - 1].Substring(rearrangeSettings.IXDESC.RearrangeColumnStart, rearrangeSettings.IXDESC.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            IXDESC = lofLines[rearrangeSettings.IXDESC.RearrangeRowNumber - 1].Substring(rearrangeSettings.IXDESC.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string IXDESC2 = "";
                    try
                    {
                        IXDESC2 = lofLines[rearrangeSettings.IXDESC2.RearrangeRowNumber - 1].Substring(rearrangeSettings.IXDESC2.RearrangeColumnStart, rearrangeSettings.IXDESC2.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            IXDESC2 = lofLines[rearrangeSettings.IXDESC2.RearrangeRowNumber - 1].Substring(rearrangeSettings.IXDESC2.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string IFENO = "";
                    try
                    {
                        IFENO = lofLines[rearrangeSettings.IFENO.RearrangeRowNumber - 1].Substring(rearrangeSettings.IFENO.RearrangeColumnStart, rearrangeSettings.IFENO.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            IFENO = lofLines[rearrangeSettings.IFENO.RearrangeRowNumber - 1].Substring(rearrangeSettings.IFENO.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Quantity = "";
                    try
                    {
                        Quantity = lofLines[rearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(rearrangeSettings.Quantity.RearrangeColumnStart, rearrangeSettings.Quantity.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Quantity = lofLines[rearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(rearrangeSettings.Quantity.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    try
                    {
                        Quantity = Quantity.Replace(".00000", "");
                        Quantity = Quantity.Replace(".0000", "");
                        Quantity = Quantity.Replace(".000", "");
                        Quantity = Quantity.Replace(".00", "");
                        Quantity = Quantity.Replace(".0", "");
                        decimal tempDecimalQuantity = Convert.ToDecimal(Quantity);
                        Quantity = tempDecimalQuantity.ToString();
                        int tempIntQuantity = Convert.ToInt32(tempDecimalQuantity);
                        Quantity = tempIntQuantity.ToString();
                    }
                    catch
                    {

                    }
                    string DateTime_Now = "";
                    try
                    {
                        DateTime_Now = lofLines[rearrangeSettings.DateTime_Now.RearrangeRowNumber - 1].Substring(rearrangeSettings.DateTime_Now.RearrangeColumnStart, rearrangeSettings.DateTime_Now.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            DateTime_Now = lofLines[rearrangeSettings.DateTime_Now.RearrangeRowNumber - 1].Substring(rearrangeSettings.DateTime_Now.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Batch = "";
                    try
                    {
                        Batch = lofLines[rearrangeSettings.Batch.RearrangeRowNumber - 1].Substring(rearrangeSettings.Batch.RearrangeColumnStart, rearrangeSettings.Batch.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Batch = lofLines[rearrangeSettings.Batch.RearrangeRowNumber - 1].Substring(rearrangeSettings.Batch.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Printer_IP = "";
                    try
                    {
                        Printer_IP = lofLines[rearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(rearrangeSettings.Printer_IP.RearrangeColumnStart, rearrangeSettings.Printer_IP.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Printer_IP = lofLines[rearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(rearrangeSettings.Printer_IP.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string NumberOfCopies = "";
                    try
                    {
                        NumberOfCopies = lofLines[rearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(rearrangeSettings.NumberOfCopies.RearrangeColumnStart, rearrangeSettings.NumberOfCopies.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            NumberOfCopies = lofLines[rearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(rearrangeSettings.NumberOfCopies.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\704-LX Template.docx";

                        if (File.Exists(originalTemplateWordDocument))
                        {
                            string IXITEM_Cust_item_number, EIX_IXDESC, EIX_IXDES2, IIM_IFENO, Quantitys, Batch_Num, DateTime_now;

                            IXITEM_Cust_item_number = "IXTEMCust";
                            EIX_IXDESC = "EIXIXDESC";
                            EIX_IXDES2 = "EIXIXD";
                            IIM_IFENO = "IIMIFENO";
                            Quantitys = "Qt";
                            Batch_Num = "Bch";
                            DateTime_now = "ToDate";

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Found Template");
                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "704-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Opened New Template");
                                documents.Replace(IXITEM_Cust_item_number, IXITEM, false, true);
                                documents.Replace(EIX_IXDESC, IXDESC, false, true);
                                documents.Replace(EIX_IXDES2, IXDESC2, false, true);
                                documents.Replace(IIM_IFENO, IFENO, false, true);
                                documents.Replace(Quantitys, Quantity, false, true);
                                documents.Replace(Batch_Num, Batch, false, true);
                                documents.Replace(DateTime_now, DateTime_Now, false, true);
                                documents.Replace(" ` ", " ", false, true);
                                documents.Save(wordTemplate);
                                documents.Close();
                            }
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "704-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Start Barcode Insert");
                                        PdfPage page = pdfDocument.Pages[0];

                                        PdfCode39Barcode barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 19;
                                        barcode.Text = IXITEM;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Size = new SizeF(120, 19);
                                        barcode.Draw(page, new PointF(13, 30));

                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Inserted Barcode: " + IXITEM);
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "704-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
                            File.Copy(newPDFFileName, outputPDFFile);
                            //PdfLoadedDocument loadedDocument = new PdfLoadedDocument(newPDFFileName);
                            //PdfDocument document = new PdfDocument();
                            //for (int i = 1; i <= totalNumberOdPages; i++)
                            //{
                            //    document.ImportPageRange(loadedDocument, 0, 0);
                            //}
                            //document.PageSettings.Orientation = PdfPageOrientation.Portrait;
                            //document.Save(outputPDFFile);
                            //document.Close(true);
                            //loadedDocument.Close(true);


                            File.Delete(wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "704-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "704-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "704-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }
    }

    public class RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LabelReareangeSettings
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

        public RearrangeSetting IXITEM
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 12
                };
            }
        }

        public RearrangeSetting IXDESC
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 67,
                    RearrangeRowNumber = 16
                };
            }
        }

        public RearrangeSetting IXDESC2
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 30,
                    RearrangeRowNumber = 15
                };
            }
        }

        public RearrangeSetting IFENO
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 37,
                    RearrangeRowNumber = 17
                };
            }
        }

        public RearrangeSetting Quantity
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 42,
                    RearrangeRowNumber = 20
                };
            }
        }

        public RearrangeSetting DateTime_Now
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 37,
                    RearrangeRowNumber = 22
                };
            }
        }

        public RearrangeSetting Batch
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 21
                };
            }
        }
    }
}
