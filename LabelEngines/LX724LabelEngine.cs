using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Syncfusion.Pdf.Barcode;
using System.Drawing;
using Syncfusion.Pdf;
using System.Net.Sockets;
using System.Net;
using System.Management;
using System.Linq;
using System.Configuration;
using Syncfusion.Pdf.Parsing;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace BarcodeLabelSoftware
{
    public class LX724LabelEngine : IDisposable
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
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
            LabelReareangeSettings rearrangeSettings = new LabelReareangeSettings();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                    string Reference_Number = "";
                    try
                    {
                        Reference_Number = lofLines[rearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.Reference_Number.RearrangeColumnStart, rearrangeSettings.Reference_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Reference_Number = lofLines[rearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.Reference_Number.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Stocking_UOM = "";
                    try
                    {
                        Stocking_UOM = lofLines[rearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(rearrangeSettings.Stocking_UOM.RearrangeColumnStart, rearrangeSettings.Stocking_UOM.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Stocking_UOM = lofLines[rearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(rearrangeSettings.Stocking_UOM.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Item_Description = "";
                    try
                    {
                        Item_Description = lofLines[rearrangeSettings.Item_Description.RearrangeRowNumber - 1].Substring(rearrangeSettings.Item_Description.RearrangeColumnStart, rearrangeSettings.Item_Description.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Description = lofLines[rearrangeSettings.Item_Description.RearrangeRowNumber - 1].Substring(rearrangeSettings.Item_Description.RearrangeColumnStart).Trim();
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
                        decimal tempDecimalQuantity = Convert.ToDecimal(Quantity);
                        int tempIntQuantity = Convert.ToInt32(tempDecimalQuantity);
                        Quantity = tempIntQuantity.ToString();
                    }
                    catch
                    {

                    }

                    string Item_Number = "";
                    try
                    {
                        Item_Number = lofLines[rearrangeSettings.Item_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.Item_Number.RearrangeColumnStart, rearrangeSettings.Item_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number = lofLines[rearrangeSettings.Item_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.Item_Number.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }


                    string To_Warehouse = "";
                    try
                    {
                        To_Warehouse = lofLines[rearrangeSettings.To_Warehouse.RearrangeRowNumber - 1].Substring(rearrangeSettings.To_Warehouse.RearrangeColumnStart, rearrangeSettings.To_Warehouse.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            To_Warehouse = lofLines[rearrangeSettings.To_Warehouse.RearrangeRowNumber - 1].Substring(rearrangeSettings.To_Warehouse.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string User = "";
                    try
                    {
                        User = lofLines[rearrangeSettings.User.RearrangeRowNumber - 1].Substring(rearrangeSettings.User.RearrangeColumnStart, rearrangeSettings.User.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            User = lofLines[rearrangeSettings.User.RearrangeRowNumber - 1].Substring(rearrangeSettings.User.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }


                    string From_Location = "";
                    try
                    {
                        From_Location = lofLines[rearrangeSettings.From_Location.RearrangeRowNumber - 1].Substring(rearrangeSettings.From_Location.RearrangeColumnStart, rearrangeSettings.From_Location.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            From_Location = lofLines[rearrangeSettings.From_Location.RearrangeRowNumber - 1].Substring(rearrangeSettings.From_Location.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }



                    string To_Location = "";
                    try
                    {
                        To_Location = lofLines[rearrangeSettings.To_Location.RearrangeRowNumber - 1].Substring(rearrangeSettings.To_Location.RearrangeColumnStart, rearrangeSettings.To_Location.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            To_Location = lofLines[rearrangeSettings.To_Location.RearrangeRowNumber - 1].Substring(rearrangeSettings.To_Location.RearrangeColumnStart).Trim();
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\724-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
                            string txtItem_Number, txtReference_Number, txtTo_Warehouse, txtFrom_Location, txtTo_Location, txtStocking_UOM, txtItem_Description, txtUser, txtQuantity, txtToDate, txtToTime;

                            txtItem_Number = "ItemNum";
                            txtReference_Number = "Ref";
                            txtTo_Warehouse = "WHS";
                            txtFrom_Location = "FrmLoc";
                            txtTo_Location = "ToLoc";
                            txtStocking_UOM = "UOMS";
                            txtItem_Description = "ItemDescription";
                            txtUser = "User";
                            txtQuantity = "Qty";
                            txtToDate = "ToDate";
                            txtToTime = "ToTime";
                            DateTime datetime = DateTime.Now;

                            try
                            {
                                decimal qty = Convert.ToDecimal(Quantity);
                                Quantity = qty.ToString();
                                int qty2 = Convert.ToInt32(qty);
                                Quantity = qty2.ToString();
                            }
                            catch
                            {

                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Found Template");
                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "724-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Opened New Template");



                                documents.Replace(txtItem_Number, Item_Number, false, true);
                                documents.Replace(txtReference_Number, Reference_Number, false, true);
                                documents.Replace(txtTo_Warehouse, To_Warehouse, false, true);
                                documents.Replace(txtFrom_Location, From_Location, false, true);
                                documents.Replace(txtTo_Location, To_Location, false, true);
                                documents.Replace(txtStocking_UOM, Stocking_UOM, false, true);
                                documents.Replace(txtItem_Description, Item_Description, false, true);
                                documents.Replace(txtUser, User, false, true);
                                documents.Replace(txtQuantity, Quantity, false, true);
                                documents.Replace(txtToDate, datetime.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(txtToTime, datetime.ToString("HH:mm:ss"), false, true);

                                documents.Save(wordTemplate);
                                documents.Close();
                            }


                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "724-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Start Barcode Insert");

                                        PdfPage page = pdfDocument.Pages[0];
                                        PdfCode39Barcode barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 16;
                                        barcode.Text = To_Warehouse;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(135, 70));

                                        PdfCode39Barcode barcode1 = new PdfCode39Barcode();
                                        barcode1.BarHeight = 16;
                                        barcode1.Text = From_Location;
                                        barcode1.TextDisplayLocation = TextLocation.None;
                                        barcode1.Draw(page, new PointF(225, 70));


                                        PdfCode39Barcode barcode2 = new PdfCode39Barcode();
                                        barcode2.BarHeight = 22;
                                        barcode2.Size = new SizeF(230, 18);
                                        barcode2.Text = Item_Number;
                                        barcode2.TextDisplayLocation = TextLocation.None;
                                        barcode2.Draw(page, new PointF(35, 150));

                                        PdfCode39Barcode barcode3 = new PdfCode39Barcode();
                                        barcode3.BarHeight = 16;
                                        barcode3.Text = Quantity;
                                        barcode3.TextDisplayLocation = TextLocation.None;
                                        barcode3.Draw(page, new PointF(35, 225));

                                        PdfCode39Barcode barcode4 = new PdfCode39Barcode();
                                        barcode4.BarHeight = 16;
                                        barcode4.Text = To_Location;
                                        barcode4.TextDisplayLocation = TextLocation.None;
                                        barcode4.Draw(page, new PointF(180, 225));


                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "724-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "724-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "724-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "724-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

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


            public RearrangeSetting Item_Number
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 37,
                        RearrangeRowNumber = 5
                    };
                }
            }


            public RearrangeSetting Reference_Number
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 37,
                        RearrangeRowNumber = 6
                    };
                }
            }
            public RearrangeSetting To_Warehouse
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 31,
                        RearrangeRowNumber = 8
                    };
                }
            }

            public RearrangeSetting From_Location
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 41,
                        RearrangeRowNumber = 9
                    };
                }
            }
            public RearrangeSetting To_Location
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 41,
                        RearrangeRowNumber = 10
                    };
                }
            }
            public RearrangeSetting Stocking_UOM
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 29,
                        RearrangeRowNumber = 12
                    };
                }
            }
            public RearrangeSetting Item_Description
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 58,
                        RearrangeRowNumber = 13
                    };
                }
            }
            public RearrangeSetting Quantity
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 48,
                        RearrangeRowNumber = 15
                    };
                }
            }

            public RearrangeSetting User

            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 36,
                        RearrangeRowNumber = 14
                    };
                }
            }
        }
    }
}
