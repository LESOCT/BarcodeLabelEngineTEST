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
    public class LX703LabelEngine : IDisposable
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
            LX703LabelReareangeSettings LX703RearrangeSettings = new LX703LabelReareangeSettings();
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);




                    // ===================== IXITEM Cust item number =====================
                    string Customer_Item_Number = "";
                    try
                    {
                        Customer_Item_Number = lofLines[LX703RearrangeSettings.Customer_Item_Number.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.Customer_Item_Number.RearrangeColumnStart, LX703RearrangeSettings.Customer_Item_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Customer_Item_Number = lofLines[LX703RearrangeSettings.Customer_Item_Number.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.Customer_Item_Number.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }




                    // ===================== IIM.IFENO Engnrng Chg lvl =====================


                    string IIM_IFENO = "";
                    try
                    {
                        IIM_IFENO = lofLines[LX703RearrangeSettings.IIM_IFENO.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.IIM_IFENO.RearrangeColumnStart, LX703RearrangeSettings.IIM_IFENO.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            IIM_IFENO = lofLines[LX703RearrangeSettings.IIM_IFENO.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.IIM_IFENO.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    // ===================== EIX.IXDESC =====================


                    string IXDESC = "";
                    try
                    {
                        IXDESC = lofLines[LX703RearrangeSettings.IXDESC.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.IXDESC.RearrangeColumnStart, LX703RearrangeSettings.IXDESC.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            IXDESC = lofLines[LX703RearrangeSettings.IXDESC.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.IXDESC.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }





                    // ===================== IIIM.IGLNO  MAT =====================


                    string MAT = "";
                    try
                    {
                        MAT = lofLines[LX703RearrangeSettings.MAT.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.MAT.RearrangeColumnStart, LX703RearrangeSettings.MAT.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            MAT = lofLines[LX703RearrangeSettings.MAT.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.MAT.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    //===================== IIM.IDSCE Colr code VW =====================

                    string ClrCode = "";
                    try
                    {
                        ClrCode = lofLines[LX703RearrangeSettings.ClrCode.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.ClrCode.RearrangeColumnStart, LX703RearrangeSettings.ClrCode.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            ClrCode = lofLines[LX703RearrangeSettings.ClrCode.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.ClrCode.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }



                    //===================== Quantity =====================


                    string Quantity = "";
                    try
                    {
                        Quantity = lofLines[LX703RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.Quantity.RearrangeColumnStart, LX703RearrangeSettings.Quantity.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Quantity = lofLines[LX703RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.Quantity.RearrangeColumnStart).TrimStart();
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
                    string Printer_IP = "";
                    try
                    {
                        Printer_IP = lofLines[LX703RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.Printer_IP.RearrangeColumnStart, LX703RearrangeSettings.Printer_IP.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Printer_IP = lofLines[LX703RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.Printer_IP.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string NumberOfCopies = "";
                    try
                    {
                        NumberOfCopies = lofLines[LX703RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.NumberOfCopies.RearrangeColumnStart, LX703RearrangeSettings.NumberOfCopies.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            NumberOfCopies = lofLines[LX703RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX703RearrangeSettings.NumberOfCopies.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\703-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Found Template");
                            string txtIIM_IFENO, txtMAT, txtIXDESC, txtCustomer_Item_Number, txtToDate, txtClrCode, txtQuantity;
                            DateTime D = DateTime.Now;


                            txtCustomer_Item_Number = "CustNum";
                            txtIXDESC = "IXDES2";
                            txtMAT = "MATs";
                            txtClrCode = "Clr_Code";
                            txtQuantity = "Qtys";
                            txtIIM_IFENO = "IFENO";
                            txtToDate = "ToDate";



                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "703-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                documents.Replace(txtCustomer_Item_Number, Customer_Item_Number, false, true);
                                documents.Replace(txtIIM_IFENO, IIM_IFENO, false, true);
                                documents.Replace(txtMAT, MAT, false, true);
                                documents.Replace(txtClrCode, ClrCode, false, true);
                                documents.Replace(txtIXDESC, IXDESC, false, true);
                                documents.Replace(txtToDate, D.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(txtQuantity, Quantity, false, true);
                                documents.Save(wordTemplate);
                                documents.Close();
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "703-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Start Barcode Insert");
                                        PdfPage page = pdfDocument.Pages[0];

                                        PdfCode39Barcode barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 15;
                                        barcode.Size = new SizeF(138, 14);
                                        barcode.Text = Customer_Item_Number;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(67, 10));

                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Failed to Insert Barcode - Error " + ex.ToString());
                                    }
                                    pdfDocument.PageSettings.Rotate = PdfPageRotateAngle.RotateAngle90;
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "703-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Deleted: " + newPDFFileName);

                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "703-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "703-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "703-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }
    }

    public class LX703RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LX703LabelReareangeSettings
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

        public LX703RearrangeSetting Customer_Item_Number
        {
            get
            {
                return new LX703RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 45,
                    RearrangeRowNumber = 12
                };
            }
        }



        public LX703RearrangeSetting IXDESC
        {
            get
            {
                return new LX703RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 60,
                    RearrangeRowNumber = 14
                };
            }
        }





        public LX703RearrangeSetting IIM_IFENO
        {
            get
            {
                return new LX703RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 36,
                    RearrangeRowNumber = 17
                };
            }
        }
        public LX703RearrangeSetting MAT
        {
            get
            {
                return new LX703RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 53,
                    RearrangeRowNumber = 18
                };
            }
        }

        // ======= NEEDS TO BE CHANGED IN TO A INT ===============
        public LX703RearrangeSetting Quantity
        {
            get
            {
                return new LX703RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 30,
                    RearrangeRowNumber = 20
                };
            }
        }

        public LX703RearrangeSetting ToDates
        {
            get
            {
                return new LX703RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 22
                };
            }
        }

        public LX703RearrangeSetting ClrCode
        {
            get
            {
                return new LX703RearrangeSetting()
                {
                    RearrangeColumnStart = 50,
                    RearrangeColumnEnd = 57,
                    RearrangeRowNumber = 25
                };
            }
        }


    }
}
