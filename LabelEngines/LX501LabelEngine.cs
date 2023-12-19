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

    namespace DocHelper.LabelEngines._501_LX
    {
        public class LX501LabelEngine : IDisposable
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
                LX501LabelReareangeSettings LX501RearrangeSettings = new LX501LabelReareangeSettings();
                LogEngine logEngine = new LogEngine();
                FileEngine csFileInputEngine = new FileEngine();
                while (lofFileData.Count > 0)
                {
                    try
                    {
                        string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);



                        string Item_Number = "";
                        try
                        {
                            Item_Number = lofLines[LX501RearrangeSettings.Item_Number.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.Item_Number.RearrangeColumnStart, LX501RearrangeSettings.Item_Number.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {
                            try
                            {
                                Item_Number = lofLines[LX501RearrangeSettings.Item_Number.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.Item_Number.RearrangeColumnStart).TrimStart();
                            }
                            catch
                            {

                            }

                        }


                        string Item_Number2 = "";
                        try
                        {
                            Item_Number2 = lofLines[LX501RearrangeSettings.Item_Number2.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.Item_Number2.RearrangeColumnStart, LX501RearrangeSettings.Item_Number2.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {
                            try
                            {
                                Item_Number2 = lofLines[LX501RearrangeSettings.Item_Number2.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.Item_Number2.RearrangeColumnStart).Trim();
                            }
                            catch
                            {

                            }
                        }



                        string Warehouse = "";
                        try
                        {
                            Warehouse = lofLines[LX501RearrangeSettings.Warehouse.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.Warehouse.RearrangeColumnStart, LX501RearrangeSettings.Warehouse.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {
                            try
                            {
                                Warehouse = lofLines[LX501RearrangeSettings.Warehouse.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.Warehouse.RearrangeColumnStart).TrimStart();
                            }
                            catch
                            {

                            }
                        }







                        string OrderNumber = "";
                        try
                        {
                            OrderNumber = lofLines[LX501RearrangeSettings.OrderNumber.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.OrderNumber.RearrangeColumnStart, LX501RearrangeSettings.OrderNumber.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                OrderNumber = lofLines[LX501RearrangeSettings.OrderNumber.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.OrderNumber.RearrangeColumnStart).Trim();
                            }
                            catch
                            {

                            }
                        }




                        string DueDate = "";
                        try
                        {
                            DueDate = lofLines[LX501RearrangeSettings.DueDate.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.DueDate.RearrangeColumnStart, LX501RearrangeSettings.DueDate.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {
                            try
                            {
                                DueDate = lofLines[LX501RearrangeSettings.DueDate.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.DueDate.RearrangeColumnStart).TrimStart();
                            }
                            catch
                            {

                            }
                        }

                        string Printer_IP = "";
                        try
                        {
                            Printer_IP = lofLines[LX501RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.Printer_IP.RearrangeColumnStart, LX501RearrangeSettings.Printer_IP.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {
                            try
                            {
                                Printer_IP = lofLines[LX501RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX501RearrangeSettings.Printer_IP.RearrangeColumnStart).Trim();
                            }
                            catch
                            {

                            }
                        }


                        try
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Start New Label");
                            string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\501 -LX Template.docx";
                            if (File.Exists(originalTemplateWordDocument))
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Found Template");

                                string txtItem_Number2, txtWarehouse, txtItem_Number, txtDueDate, txtOrderNumber;
                                DateTime D = DateTime.Now;


                                txtItem_Number = "ItemNum";
                                txtItem_Number2 = "ItemNum2";
                                txtWarehouse = "WHS";
                                txtOrderNumber = "OrderNum";
                                txtDueDate = "Duedate";


                                string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "501 -LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                                File.Copy(originalTemplateWordDocument, wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "New Work Template Created: " + wordTemplate);
                                using (WordDocument documents = new WordDocument())
                                {
                                    documents.Open(wordTemplate);
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Opened New Template");


                                    documents.Replace(txtItem_Number, Item_Number, false, true);
                                    documents.Replace(txtItem_Number2, Item_Number2, false, true);
                                    documents.Replace(txtOrderNumber, OrderNumber, false, true);
                                    documents.Replace(txtWarehouse, Warehouse, false, true);
                                    documents.Replace(txtDueDate, DueDate, false, true);



                                    // ======== QUANTITY DOESNT WANT TO CONVERT 
                                    documents.Save(wordTemplate);
                                    documents.Close();
                                }


                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Saved and Closed Template");

                                string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "501 -LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "New PDF Document Created: " + newPDFFileName);
                                using (DocToPDFConverter converter = new DocToPDFConverter())
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                    using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Converted and Saved PDF Document");

                                        try
                                        {
                                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Start Barcode Insert");

                                            PdfPage page = pdfDocument.Pages[0];

                                            PdfCode39Barcode barcode1 = new PdfCode39Barcode();

                                            barcode1.BarHeight = 18;
                                            barcode1.Text = Item_Number;
                                            barcode1.TextDisplayLocation = TextLocation.None;
                                            barcode1.Draw(page, new PointF(12, 7));

                                            PdfCode39Barcode barcode = new PdfCode39Barcode();

                                            barcode.BarHeight = 18;
                                            barcode.Text = OrderNumber;
                                            barcode.TextDisplayLocation = TextLocation.Bottom;
                                            barcode.Draw(page, new PointF(100, 85));


                                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Inserted Barcode: " + Item_Number);
                                        }
                                        catch (Exception ex)
                                        {
                                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Failed to Insert Barcode - Error " + ex.ToString());
                                        }

                                        pdfDocument.Save(newPDFFileName);
                                        pdfDocument.Close(true);
                                    }
                                }

                                int totalNumberOdPages = 1;
                                try
                                {
                                    totalNumberOdPages = Convert.ToInt32(2);
                                }
                                catch
                                {

                                }

                                string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "501 -LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Deleted: " + wordTemplate);
                                File.Delete(newPDFFileName);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Deleted: " + newPDFFileName);
                            }
                            else
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "No Template Found");
                            }
                        }
                        catch (Exception ex)
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Failed to Process - Error " + ex.ToString());
                        }

                        try
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX", "Finished Processing Label");
                            csFileInputEngine.MoveFileToArchive(lofFileData[0], "501-LX");
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "501-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                            csFileInputEngine.MoveFileToArchive(lofFileData[0], "501-LX FAILED");
                            lofFileData.RemoveAt(0);
                        }
                        catch
                        {

                        }
                    }
                }
            }

            public class LX501RearrangeSetting
            {
                public int RearrangeColumnStart { get; set; }
                public int RearrangeColumnEnd { get; set; }
                public int RearrangeRowNumber { get; set; }
            }

            public class LX501LabelReareangeSettings
            {

                public LX501RearrangeSetting Item_Number
                {
                    get
                    {
                        return new LX501RearrangeSetting()
                        {
                            RearrangeColumnStart = 5,
                            RearrangeColumnEnd = 22,
                            RearrangeRowNumber = 9
                        };
                    }
                }

                public LX501RearrangeSetting Item_Number2
                {
                    get
                    {
                        return new LX501RearrangeSetting()
                        {
                            RearrangeColumnStart = 5,
                            RearrangeColumnEnd = 28,
                            RearrangeRowNumber = 10
                        };
                    }
                }

                public LX501RearrangeSetting Warehouse
                {
                    get
                    {
                        return new LX501RearrangeSetting()
                        {
                            RearrangeColumnStart = 38,
                            RearrangeColumnEnd = 2,
                            RearrangeRowNumber = 12
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





                // ======= NEEDS TO BE CHANGED IN TO A INT ===============

                public LX501RearrangeSetting OrderNumber
                {
                    get
                    {
                        return new LX501RearrangeSetting()
                        {
                            RearrangeColumnStart = 26,
                            RearrangeColumnEnd = 7,
                            RearrangeRowNumber = 6
                        };
                    }
                }


                public LX501RearrangeSetting DueDate
                {
                    get
                    {
                        return new LX501RearrangeSetting()
                        {
                            RearrangeColumnStart = 51,
                            RearrangeColumnEnd = 7,
                            RearrangeRowNumber = 6
                        };
                    }
                }

            }
        }
    }
}
