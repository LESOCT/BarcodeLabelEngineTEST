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
    public class LX712LabelEngine : IDisposable
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

                    string PO_Number = "";
                    try
                    {
                        PO_Number = lofLines[rearrangeSettings.PO_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.PO_Number.RearrangeColumnStart, rearrangeSettings.PO_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            PO_Number = lofLines[rearrangeSettings.PO_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.PO_Number.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
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
                    string PO_Line_Number = "";
                    try
                    {
                        PO_Line_Number = lofLines[rearrangeSettings.PO_Line_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.PO_Line_Number.RearrangeColumnStart, rearrangeSettings.PO_Line_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            PO_Line_Number = lofLines[rearrangeSettings.PO_Line_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.PO_Line_Number.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Warehouse = "";
                    try
                    {
                        Warehouse = lofLines[rearrangeSettings.Warehouse.RearrangeRowNumber - 1].Substring(rearrangeSettings.Warehouse.RearrangeColumnStart, rearrangeSettings.Warehouse.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Warehouse = lofLines[rearrangeSettings.Warehouse.RearrangeRowNumber - 1].Substring(rearrangeSettings.Warehouse.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Lot = "";
                    try
                    {
                        Lot = lofLines[rearrangeSettings.Lot.RearrangeRowNumber - 1].Substring(rearrangeSettings.Lot.RearrangeColumnStart, rearrangeSettings.Lot.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Lot = lofLines[rearrangeSettings.Lot.RearrangeRowNumber - 1].Substring(rearrangeSettings.Lot.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }


                    string Location = "";
                    try
                    {
                        Location = lofLines[rearrangeSettings.Location.RearrangeRowNumber - 1].Substring(rearrangeSettings.Location.RearrangeColumnStart, rearrangeSettings.Location.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Location = lofLines[rearrangeSettings.Location.RearrangeRowNumber - 1].Substring(rearrangeSettings.Location.RearrangeColumnStart).Trim();
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
                    string Quantity = "";
                    try
                    {
                        Quantity = lofLines[rearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(rearrangeSettings.Quantity.RearrangeColumnStart, rearrangeSettings.Quantity.RearrangeColumnEnd).Trim();
                        try { Quantity = Quantity.Replace(".00", "").Replace(".0", ""); decimal IntQuantity = Convert.ToDecimal(Quantity); Quantity = Convert.ToInt32(IntQuantity).ToString(); } catch { }
                    }
                    catch
                    {
                        Quantity = lofLines[rearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(rearrangeSettings.Quantity.RearrangeColumnStart).Trim();
                        try { Quantity = Quantity.Replace(".00", "").Replace(".0", ""); decimal IntQuantity = Convert.ToDecimal(Quantity); Quantity = Convert.ToInt32(IntQuantity).ToString(); } catch { }
                    }

                    string Garangua_Label = "";
                    try
                    {
                        Garangua_Label = lofLines[rearrangeSettings.Garangua_Label.RearrangeRowNumber - 1].Substring(rearrangeSettings.Garangua_Label.RearrangeColumnStart, rearrangeSettings.Garangua_Label.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Garangua_Label = lofLines[rearrangeSettings.Garangua_Label.RearrangeRowNumber - 1].Substring(rearrangeSettings.Garangua_Label.RearrangeColumnStart).Trim();
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

                    if (Warehouse == "V1")
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Did not Processing Label for " + Warehouse);
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "712-LX Discarded" + Warehouse);
                        lofFileData.RemoveAt(0);
                    }
                    else
                    {
                        try
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Start New Label");
                            string originalTemplateWordDocument = "";
                            if (Warehouse == "U1" || Warehouse == "U9" || Warehouse == "U2" || Warehouse == "U8")
                            {
                                originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\712A-LX Template.docx";
                            }
                            else
                            {
                                originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\712B-LX Template.docx";
                            }

                            if (File.Exists(originalTemplateWordDocument))
                            {
                                if (Printer_IP == "")
                                {
                                    Printer_IP = "192.168.5.204";
                                }

                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Found Template");
                                string txtItem_Number, txtPO_Number, txtWarehouse, txtLot, txtGarangua_Label, txtPO_Line_Number, txtLocation, txtStocking_UOM, txtItem_Description, txtUser, txtQuantity, txtToDate, txtToTime;

                                txtPO_Number = "PO_order";
                                txtItem_Number = "ItemNum";
                                txtWarehouse = "WS";
                                txtPO_Line_Number = "PO_Line";
                                txtLocation = "Location";
                                txtStocking_UOM = "uomS";
                                txtItem_Description = "ItemDescription";
                                txtUser = "Users";
                                txtQuantity = "Qtys";
                                txtGarangua_Label = "Garlabel";
                                txtLot = "Lot1";
                                txtToDate = "ToDate";
                                txtToTime = "ToTime";

                                DateTime datetime = DateTime.Now;



                                string wordTemplate = "";
                                if (Warehouse == "U1" || Warehouse == "U9" || Warehouse == "U2" || Warehouse == "U8")
                                {
                                    wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "712A-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                                }
                                else
                                {
                                    wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "712B-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                                }


                                File.Copy(originalTemplateWordDocument, wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "New Work Template Created: " + wordTemplate);
                                using (WordDocument documents = new WordDocument())
                                {
                                    documents.Open(wordTemplate);
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Opened New Template");

                                    documents.Replace(txtPO_Number, PO_Number, false, true);
                                    documents.Replace(txtItem_Number, Item_Number, false, true);
                                    documents.Replace(txtPO_Line_Number, PO_Line_Number, false, true);
                                    documents.Replace(txtWarehouse, Warehouse, false, true);
                                    documents.Replace(txtLocation, Location, false, true);
                                    documents.Replace(txtLot, Lot, false, true);
                                    documents.Replace(txtStocking_UOM, Stocking_UOM, false, true);
                                    documents.Replace(txtItem_Description, Item_Description, false, true);
                                    documents.Replace(txtUser, User, false, true);
                                    documents.Replace(txtQuantity, Quantity, false, true);
                                    documents.Replace(txtToDate, datetime.ToString("dd/MM/yyyy"), false, true);
                                    documents.Replace(txtToTime, datetime.ToString("HH:mm:ss"), false, true);
                                    documents.Replace(txtGarangua_Label, Garangua_Label, false, true);

                                    if (Warehouse == "U1" || Warehouse == "U9" || Warehouse == "U2" || Warehouse == "U8") { }
                                    else
                                    {
                                        documents.Replace("ScanDate", "FIFO " + datetime.ToString("MMMM"), false, true);
                                    }

                                    documents.Save(wordTemplate);
                                    documents.Close();
                                }

                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Saved and Closed Template");

                                string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "712-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "New PDF Document Created: " + newPDFFileName);
                                using (DocToPDFConverter converter = new DocToPDFConverter())
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                    using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Converted and Saved PDF Document");

                                        try
                                        {
                                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Start Barcode Insert");
                                            PdfPage page = pdfDocument.Pages[0];

                                            if (Warehouse == "U1" || Warehouse == "U9" || Warehouse == "U2" || Warehouse == "U8")
                                            {
                                                PdfCode39Barcode barcode = new PdfCode39Barcode();
                                                barcode.BarHeight = 18;
                                                barcode.Text = Warehouse;
                                                barcode.TextDisplayLocation = TextLocation.None;
                                                barcode.Draw(page, new PointF(100, 60));

                                                PdfCode39Barcode barcode1 = new PdfCode39Barcode();
                                                barcode1.BarHeight = 18;
                                                barcode1.Text = Location;
                                                barcode1.Size = new SizeF(85, 18);
                                                barcode1.TextDisplayLocation = TextLocation.None;
                                                barcode1.Draw(page, new PointF(170, 60));

                                                PdfCode39Barcode barcode3 = new PdfCode39Barcode();
                                                barcode3.BarHeight = 20;
                                                barcode3.Text = Item_Number;
                                                barcode3.Size = new SizeF(165, 20);
                                                barcode3.TextDisplayLocation = TextLocation.None;
                                                barcode3.Draw(page, new PointF(20, 117));

                                                PdfCode39Barcode barcode4 = new PdfCode39Barcode();
                                                barcode4.BarHeight = 18;
                                                barcode4.Text = Quantity;
                                                barcode4.TextDisplayLocation = TextLocation.None;
                                                barcode4.Draw(page, new PointF(20, 212));
                                            }
                                            else
                                            {
                                                PdfCode39Barcode barcode = new PdfCode39Barcode();
                                                barcode.BarHeight = 15;
                                                barcode.Text = Warehouse;
                                                barcode.TextDisplayLocation = TextLocation.None;
                                                barcode.Draw(page, new PointF(130, 70));

                                                PdfCode39Barcode barcode1 = new PdfCode39Barcode();
                                                barcode1.BarHeight = 15;
                                                barcode1.Text = Location;
                                                barcode1.TextDisplayLocation = TextLocation.None;
                                                barcode1.Draw(page, new PointF(210, 70));

                                                PdfCode39Barcode barcode3 = new PdfCode39Barcode();
                                                barcode3.BarHeight = 22;
                                                barcode3.Text = Item_Number;
                                                barcode3.TextDisplayLocation = TextLocation.None;
                                                barcode3.Draw(page, new PointF(30, 127));

                                                PdfCode39Barcode barcode4 = new PdfCode39Barcode();
                                                barcode4.BarHeight = 18;
                                                barcode4.Text = Quantity;
                                                barcode4.TextDisplayLocation = TextLocation.None;
                                                barcode4.Draw(page, new PointF(20, 212));

                                                if (!string.IsNullOrWhiteSpace(Lot))
                                                {
                                                    PdfCode39Barcode barcode6 = new PdfCode39Barcode();
                                                    barcode6.BarHeight = 18;
                                                    barcode6.Text = Lot;
                                                    barcode6.TextDisplayLocation = TextLocation.None;
                                                    barcode6.Draw(page, new PointF(210, 212));
                                                }
                                            }

                                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Inserted Barcodes");
                                        }
                                        catch (Exception ex)
                                        {
                                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                                string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "712-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Deleted: " + wordTemplate);
                                File.Delete(newPDFFileName);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Deleted: " + newPDFFileName);
                            }
                            else
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "No Template Found");
                            }
                        }
                        catch (Exception ex)
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Failed to Process - Error " + ex.ToString());
                        }

                        try
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX", "Finished Processing Label");
                            csFileInputEngine.MoveFileToArchive(lofFileData[0], "712-LX");
                            lofFileData.RemoveAt(0);
                        }
                        catch
                        {

                        }
                    }
                }
                catch (Exception ex)
                {
                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "712-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "712-LX FAILED");
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

            public RearrangeSetting PO_Number
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
            public RearrangeSetting Item_Number
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
            public RearrangeSetting PO_Line_Number
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 45,
                        RearrangeRowNumber = 7
                    };
                }
            }


            public RearrangeSetting Warehouse
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 32,
                        RearrangeRowNumber = 8
                    };
                }
            }

            public RearrangeSetting Location
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
            public RearrangeSetting Lot
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
                        RearrangeColumnEnd = 59,
                        RearrangeRowNumber = 14
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
                        RearrangeRowNumber = 15
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
                        RearrangeColumnEnd = 34,
                        RearrangeRowNumber = 16
                    };
                }
            }


            public RearrangeSetting Garangua_Label
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 29,
                        RearrangeRowNumber = 17
                    };
                }
            }

        }
    }
}
