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
    public class LX708LabelEngine : IDisposable
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
                LogEngine logEngine = new LogEngine();logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Tasks", "Disposed Label Engine");handle.Dispose();
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
            LX708LabelReareangeSettings LX707RearrangeSettings = new LX708LabelReareangeSettings();
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);



                    string Toyota_CKD = "";
                    try
                    {
                        Toyota_CKD = lofLines[LX707RearrangeSettings.Toyota_CKD.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Toyota_CKD.RearrangeColumnStart, LX707RearrangeSettings.Toyota_CKD.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        Toyota_CKD = lofLines[LX707RearrangeSettings.Toyota_CKD.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Toyota_CKD.RearrangeColumnStart).TrimStart();
                    }






                    string Item_Number = "";
                    try
                    {
                        Item_Number = lofLines[LX707RearrangeSettings.Item_Number.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Item_Number.RearrangeColumnStart, LX707RearrangeSettings.Item_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        Item_Number = lofLines[LX707RearrangeSettings.Item_Number.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Item_Number.RearrangeColumnStart).Trim();
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




                    string Cust_Item_Number = "";
                    try
                    {
                        Cust_Item_Number = lofLines[LX707RearrangeSettings.Cust_Item_Number.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Cust_Item_Number.RearrangeColumnStart, LX707RearrangeSettings.Cust_Item_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        Cust_Item_Number = lofLines[LX707RearrangeSettings.Cust_Item_Number.RearrangeRowNumber - 1].Substring(LX707RearrangeSettings.Cust_Item_Number.RearrangeColumnStart).Trim();
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\708-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Found Template");
                            string txtBatch, txtCustItemNum, txtToTime, txtToyota_CKD, txtToDate, txtUser;
                            DateTime D = DateTime.Now;


                            txtToyota_CKD = "C1";
                            txtBatch = "Bch";
                            txtUser = "Users";
                            txtToDate = "ToDate";
                            txtToTime = "Totime";
                            txtCustItemNum = "CustItem";





                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "708-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {

                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Opened New Template");

                                documents.Replace(txtToyota_CKD, Toyota_CKD, false, true);
                                documents.Replace(txtBatch, Batch, false, true);
                                documents.Replace(txtToDate, D.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(txtToTime, D.ToString("HH:mm:ss"), false, true);
                                documents.Replace(txtUser, User, false, true);
                                documents.Replace(txtCustItemNum, Cust_Item_Number, false, true);




                                // ======== QUANTITY DOESNT WANT TO CONVERT 



                                documents.Save(wordTemplate);
                                documents.Close();
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "708-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Converted and Saved PDF Document");


                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Start Barcode Insert");

                                        PdfPage page = pdfDocument.Pages[0];

                                        PdfCode39Barcode barcode1 = new PdfCode39Barcode();
                                        barcode1.BarHeight = 22;
                                        barcode1.Text = Item_Number;
                                        barcode1.TextDisplayLocation = TextLocation.None;
                                        barcode1.Size = new SizeF(160, 22);
                                        barcode1.Draw(page, new PointF(7, 180));



                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "708-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "708-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "708-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "708-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

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
            public LX708RearrangeSetting User
            {
                get
                {
                    return new LX708RearrangeSetting()
                    {
                        RearrangeColumnStart = 27,
                        RearrangeColumnEnd = 38,
                        RearrangeRowNumber = 19
                    };
                }
            }


            public LX708RearrangeSetting Toyota_CKD
            {
                get
                {
                    return new LX708RearrangeSetting()
                    {
                        RearrangeColumnStart = 27,
                        RearrangeColumnEnd = 29,
                        RearrangeRowNumber = 30
                    };
                }
            }

            public LX708RearrangeSetting Batch
            {
                get
                {
                    return new LX708RearrangeSetting()
                    {
                        RearrangeColumnStart = 27,
                        RearrangeColumnEnd = 40,
                        RearrangeRowNumber = 21
                    };
                }
            }
            public LX708RearrangeSetting Item_Number
            {
                get
                {
                    return new LX708RearrangeSetting()
                    {
                        RearrangeColumnStart = 27,
                        RearrangeColumnEnd = 47,
                        RearrangeRowNumber = 6
                    };
                }
            }

            public LX708RearrangeSetting Cust_Item_Number
            {
                get
                {
                    return new LX708RearrangeSetting()
                    {
                        RearrangeColumnStart = 27,
                        RearrangeColumnEnd = 40,
                        RearrangeRowNumber = 12
                    };
                }
            }




        }
    }
}
