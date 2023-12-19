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
    public class LX710LabelEngine : IDisposable
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
            LX708LabelReareangeSettings LX707RearrangeSettings = new LX708LabelReareangeSettings();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                    string PO_Number = "";
                    try
                    {
                        PO_Number = lofLines[LX707RearrangeSettings.PO_Number.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.PO_Number.RearrangeColumnStart, LX707RearrangeSettings.PO_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            PO_Number = lofLines[LX707RearrangeSettings.PO_Number.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.PO_Number.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }



                    string Item_Number = "";
                    try
                    {
                        Item_Number = lofLines[LX707RearrangeSettings.Item_Number.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Item_Number.RearrangeColumnStart, LX707RearrangeSettings.Item_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number = lofLines[LX707RearrangeSettings.Item_Number.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Item_Number.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }



                    string Stock_UOM = "";
                    try
                    {
                        Stock_UOM = lofLines[LX707RearrangeSettings.Stock_UOM.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Stock_UOM.RearrangeColumnStart, LX707RearrangeSettings.Stock_UOM.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Stock_UOM = lofLines[LX707RearrangeSettings.Stock_UOM.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Stock_UOM.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }



                    string Item_Description = "";
                    try
                    {
                        Item_Description = lofLines[LX707RearrangeSettings.Item_Description.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Item_Description.RearrangeColumnStart, LX707RearrangeSettings.Item_Description.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Description = lofLines[LX707RearrangeSettings.Item_Description.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Item_Description.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Quantity = "";
                    try
                    {
                        Quantity = lofLines[LX707RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Quantity.RearrangeColumnStart, LX707RearrangeSettings.Quantity.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Quantity = lofLines[LX707RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Quantity.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    try
                    {
                        decimal tempDecimalQuantity = Convert.ToDecimal(Quantity);
                        int tempIntQuantity = Convert.ToInt32(tempDecimalQuantity);
                        Quantity = tempIntQuantity.ToString();
                    }
                    catch
                    {

                    }

                    string Cust_Item_Description = "";
                    try
                    {
                        Cust_Item_Description = lofLines[LX707RearrangeSettings.Cust_Item_Description.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Cust_Item_Description.RearrangeColumnStart, LX707RearrangeSettings.Cust_Item_Description.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Cust_Item_Description = lofLines[LX707RearrangeSettings.Cust_Item_Description.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Cust_Item_Description.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\710-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Found Template");
                            string txtItem_Number, txtCust_Item_Description, txtPO_Number, txtToDate, txtItem_Description, txtStock_UOM, txtQuantity;
                            DateTime D = DateTime.Now;


                            txtPO_Number = "PO_Num";
                            txtItem_Number = "ItmNum";
                            txtStock_UOM = "UOMs";
                            txtItem_Description = "Item_Description";
                            txtToDate = "ToDate";
                            txtQuantity = "Qtys";
                            txtCust_Item_Description = "CustItemDesc";


                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "710-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);

                                documents.Replace(txtPO_Number, PO_Number, false, true);
                                documents.Replace(txtItem_Number, Item_Number, false, true);
                                documents.Replace(txtStock_UOM, Stock_UOM, false, true);
                                documents.Replace(txtItem_Description, Item_Description, false, true);
                                documents.Replace(txtToDate, D.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(txtQuantity, Quantity, false, true);
                                documents.Replace(txtCust_Item_Description, Cust_Item_Description, false, true);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Opened New Template");

                                // ======== QUANTITY DOESNT WANT TO CONVERT 



                                documents.Save(wordTemplate);
                                documents.Close();
                            }
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "710-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Start Barcode Insert");
                                        //PdfPage page = pdfDocument.Pages[0];

                                        try
                                        {
                                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Start Barcode Insert");
                                            PdfPage page = pdfDocument.Pages[0];

                                            PdfCode39Barcode barcode = new PdfCode39Barcode();
                                            barcode.BarHeight = 24;
                                            barcode.Text = Item_Number;
                                            barcode.TextDisplayLocation = TextLocation.None;
                                            barcode.Draw(page, new PointF(195, 130));

                                            PdfCode39Barcode barcode1 = new PdfCode39Barcode();
                                            barcode1.BarHeight = 15;
                                            barcode1.Text = Quantity;
                                            barcode1.TextDisplayLocation = TextLocation.None;
                                            barcode1.Draw(page, new PointF(210, 245));

                                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Inserted Barcodes");
                                        }
                                        catch (Exception ex)
                                        {
                                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Failed to Insert Barcode - Error " + ex.ToString());
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "710-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "710-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "710-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "710-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }
    }

    public class LX708RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LX708LabelReareangeSettings
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
        public LX708RearrangeSetting PO_Number
        {
            get
            {
                return new LX708RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 37,
                    RearrangeRowNumber = 5
                };
            }
        }

        public LX708RearrangeSetting Item_Number
        {
            get
            {
                return new LX708RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 6
                };
            }
        }



        public LX708RearrangeSetting Stock_UOM
        {
            get
            {
                return new LX708RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 30,
                    RearrangeRowNumber = 14
                };
            }
        }


        public LX708RearrangeSetting Item_Description
        {
            get
            {
                return new LX708RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 31,
                    RearrangeRowNumber = 15
                };
            }
        }

        public LX708RearrangeSetting Quantity
        {
            get
            {
                return new LX708RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 48,
                    RearrangeRowNumber = 16
                };
            }
        }
        public LX708RearrangeSetting Cust_Item_Description
        {
            get
            {
                return new LX708RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 58,
                    RearrangeRowNumber = 19
                };
            }
        }
    }
}
