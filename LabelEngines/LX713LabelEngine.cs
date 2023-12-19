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
    public class LX713LabelEngine : IDisposable
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
            LX713LabelReareangeSettings rearrangeSettings = new LX713LabelReareangeSettings();
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
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
                        IXITEM = lofLines[rearrangeSettings.IXITEM.RearrangeRowNumber - 1].Substring(rearrangeSettings.IXITEM.RearrangeColumnStart).Trim();
                    }

                    string IXDESC2 = "";
                    try
                    {
                        IXDESC2 = lofLines[rearrangeSettings.IXDESC2.RearrangeRowNumber - 1].Substring(rearrangeSettings.IXDESC2.RearrangeColumnStart, rearrangeSettings.IXDESC2.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        IXDESC2 = lofLines[rearrangeSettings.IXDESC2.RearrangeRowNumber - 1].Substring(rearrangeSettings.IXDESC2.RearrangeColumnStart).Trim();
                    }

                    string Toyota_CKD = "";
                    try
                    {
                        Toyota_CKD = lofLines[rearrangeSettings.Toyota_CKD.RearrangeRowNumber - 1].Substring(rearrangeSettings.Toyota_CKD.RearrangeColumnStart, rearrangeSettings.Toyota_CKD.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        Toyota_CKD = lofLines[rearrangeSettings.Toyota_CKD.RearrangeRowNumber - 1].Substring(rearrangeSettings.Toyota_CKD.RearrangeColumnStart).Trim();
                    }

                    string Quantity = "";
                    try
                    {
                        Quantity = lofLines[rearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(rearrangeSettings.Quantity.RearrangeColumnStart, rearrangeSettings.Quantity.RearrangeColumnEnd).Trim();
                        try { Quantity = Convert.ToInt32(Quantity).ToString(); } catch { }
                    }
                    catch
                    {
                        Quantity = lofLines[rearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(rearrangeSettings.Quantity.RearrangeColumnStart).Trim();
                        try { Quantity = Convert.ToInt32(Quantity).ToString(); } catch { }
                    }


                    string Batch = "";
                    try
                    {
                        Batch = lofLines[rearrangeSettings.Batch.RearrangeRowNumber - 1].Substring(rearrangeSettings.Batch.RearrangeColumnStart, rearrangeSettings.Batch.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        Batch = lofLines[rearrangeSettings.Batch.RearrangeRowNumber - 1].Substring(rearrangeSettings.Batch.RearrangeColumnStart).Trim();
                    }

                    string Cust_Item_Number = "";
                    try
                    {
                        Cust_Item_Number = lofLines[rearrangeSettings.Cust_Item_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.Cust_Item_Number.RearrangeColumnStart, rearrangeSettings.Cust_Item_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        Cust_Item_Number = lofLines[rearrangeSettings.Cust_Item_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.Cust_Item_Number.RearrangeColumnStart).Trim();
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\713-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
                            string txtQuantitys, txtIXITEM, Batch_Num, txtEIXDESC2, txtToyota_CKD;

                            txtIXITEM = "IXITEM";
                            txtEIXDESC2 = "IDESC2";
                            txtToyota_CKD = "ToyotaCKD";
                            txtQuantitys = "Qt";
                            Batch_Num = "Bch";

                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "713-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Opened New Template");

                                documents.Replace(txtIXITEM, IXITEM, false, true);
                                documents.Replace(txtEIXDESC2, IXDESC2, false, true);
                                documents.Replace(txtToyota_CKD, Toyota_CKD, false, true);
                                documents.Replace(txtQuantitys, Quantity, false, true);
                                documents.Replace(Batch_Num, Batch, false, true);
                                documents.Save(wordTemplate);
                                documents.Close();
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "713-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Start Barcode Insert");
                                        PdfPage page = pdfDocument.Pages[0];
                                        PdfCode39Barcode barcode1 = new PdfCode39Barcode();
                                        barcode1.BarHeight = 21;
                                        barcode1.Text = Cust_Item_Number;
                                        barcode1.TextDisplayLocation = TextLocation.None;
                                        barcode1.Size = new SizeF(170, 20);
                                        barcode1.Draw(page, new PointF(18, 70));


                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "713-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "713-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "713-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "713-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }
    }


    public class LX713RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LX713LabelReareangeSettings
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

        public RearrangeSetting Cust_Item_Number
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 53,
                    RearrangeRowNumber = 12
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
                    RearrangeColumnEnd = 33,
                    RearrangeRowNumber = 4
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
                    RearrangeColumnEnd = 53,
                    RearrangeRowNumber = 12
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
                    RearrangeColumnEnd = 37,
                    RearrangeRowNumber = 21
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
                    RearrangeColumnEnd = 61,
                    RearrangeRowNumber = 28
                };
            }
        }

        public RearrangeSetting Toyota_CKD
        {
            get
            {
                return new RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 30
                };
            }
        }


    }

}
