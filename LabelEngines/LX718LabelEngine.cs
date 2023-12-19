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
    public class LX718LabelEngine : IDisposable
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
            LX718LabelReareangeSettings LX718RearrangeSettings = new LX718LabelReareangeSettings();
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);


                    string Warehouse = "";
                    try
                    {
                        Warehouse = lofLines[LX718RearrangeSettings.Warehouse.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Warehouse.RearrangeColumnStart, LX718RearrangeSettings.Warehouse.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Warehouse = lofLines[LX718RearrangeSettings.Warehouse.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Warehouse.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Item_Number_first = "";
                    try
                    {
                        Item_Number_first = lofLines[LX718RearrangeSettings.Item_Number_first.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Item_Number_first.RearrangeColumnStart, LX718RearrangeSettings.Item_Number_first.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number_first = lofLines[LX718RearrangeSettings.Item_Number_first.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Item_Number_first.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Item_Number_second = "";
                    try
                    {
                        Item_Number_second = lofLines[LX718RearrangeSettings.Item_Number_second.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Item_Number_second.RearrangeColumnStart, LX718RearrangeSettings.Item_Number_second.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number_second = lofLines[LX718RearrangeSettings.Item_Number_second.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Item_Number_second.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Location = "";
                    try
                    {
                        Location = lofLines[LX718RearrangeSettings.Location.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Location.RearrangeColumnStart, LX718RearrangeSettings.Location.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Location = lofLines[LX718RearrangeSettings.Location.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Location.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string User = "";
                    try
                    {
                        User = lofLines[LX718RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.User.RearrangeColumnStart, LX718RearrangeSettings.User.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            User = lofLines[LX718RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.User.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Stocking_UOM = "";
                    try
                    {
                        Stocking_UOM = lofLines[LX718RearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Stocking_UOM.RearrangeColumnStart, LX718RearrangeSettings.Stocking_UOM.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Stocking_UOM = lofLines[LX718RearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Stocking_UOM.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Lot_First = "";
                    try
                    {
                        Lot_First = lofLines[LX718RearrangeSettings.Lot_First.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Lot_First.RearrangeColumnStart, LX718RearrangeSettings.Lot_First.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Lot_First = lofLines[LX718RearrangeSettings.Lot_First.RearrangeRowNumber - 2].Substring(LX718RearrangeSettings.Lot_First.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Item_description = "";
                    try
                    {
                        Item_description = lofLines[LX718RearrangeSettings.Item_description.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Item_description.RearrangeColumnStart, LX718RearrangeSettings.Item_description.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_description = lofLines[LX718RearrangeSettings.Item_description.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Item_description.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Lot_Second = "";
                    try
                    {
                        Lot_Second = lofLines[LX718RearrangeSettings.Lot_Second.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Lot_Second.RearrangeColumnStart, LX718RearrangeSettings.Lot_Second.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Lot_Second = lofLines[LX718RearrangeSettings.Lot_Second.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Lot_Second.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }
                    string Item_Type = "";
                    try
                    {
                        Item_Type = lofLines[LX718RearrangeSettings.Item_Type.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Item_Type.RearrangeColumnStart, LX718RearrangeSettings.Item_Type.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Type = lofLines[LX718RearrangeSettings.Item_Type.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Item_Type.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string Quantity = "";
                    try
                    {
                        Quantity = lofLines[LX718RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Quantity.RearrangeColumnStart, LX718RearrangeSettings.Quantity.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Quantity = lofLines[LX718RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Quantity.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    } 

                    try
                    {
                        Quantity = Quantity.Replace(".00000", ",00");
                        Quantity = Quantity.Replace(".0000", ",00");
                        Quantity = Quantity.Replace(".000", ",00");
                        Quantity = Quantity.Replace(".00", ",00");
                        Quantity = Quantity.Replace(".0", ",00");
                        decimal tempDecimalQuantity = Math.Round(Convert.ToDecimal(Quantity), 2);
                        Quantity = tempDecimalQuantity.ToString();
                        if ((tempDecimalQuantity % 1) == 0)
                        {
                            int tempIntQuantity = Convert.ToInt32(tempDecimalQuantity);
                            Quantity = tempIntQuantity.ToString();
                        }
                    }
                    catch
                    {

                    }

                    string CustNumber = "";
                    try
                    {
                        CustNumber = lofLines[LX718RearrangeSettings.CustNumber.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.CustNumber.RearrangeColumnStart, LX718RearrangeSettings.CustNumber.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            CustNumber = lofLines[LX718RearrangeSettings.CustNumber.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.CustNumber.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string Back_Number = "";
                    try
                    {
                        Back_Number = lofLines[LX718RearrangeSettings.Back_Number.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Back_Number.RearrangeColumnStart, LX718RearrangeSettings.Back_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Back_Number = lofLines[LX718RearrangeSettings.Back_Number.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Back_Number.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }
                    string Reference_Number = "";
                    try
                    {
                        Reference_Number = lofLines[LX718RearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Reference_Number.RearrangeColumnStart, LX718RearrangeSettings.Reference_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Reference_Number = lofLines[LX718RearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Reference_Number.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string ShopOrder = "";
                    try
                    {
                        ShopOrder = lofLines[LX718RearrangeSettings.ShopOrder.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.ShopOrder.RearrangeColumnStart, LX718RearrangeSettings.ShopOrder.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            ShopOrder = lofLines[LX718RearrangeSettings.ShopOrder.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.ShopOrder.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string Printer_IP = "";
                    try
                    {
                        Printer_IP = lofLines[LX718RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Printer_IP.RearrangeColumnStart, LX718RearrangeSettings.Printer_IP.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Printer_IP = lofLines[LX718RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.Printer_IP.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string NumberOfCopies = "";
                    try
                    {
                        NumberOfCopies = lofLines[LX718RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.NumberOfCopies.RearrangeColumnStart, LX718RearrangeSettings.NumberOfCopies.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            NumberOfCopies = lofLines[LX718RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX718RearrangeSettings.NumberOfCopies.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\718-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Found Template");
                            string txtWarehouse, txtReference_Number, txtLocation, txtUser, txtItem_Number_first, txtItem_Number_second, txtItem_Type, txtStocking_UOM, txtItem_description, txtBackNumber, txtQuantity, txtCustNumber, txtLot_Second, txtLot_First, DateTime_now, Shop_Order;
                            DateTime D = DateTime.Now;
                            txtItem_Number_first = "Tem";
                            txtItem_Number_second = "Item2";
                            txtWarehouse = "Whs";
                            txtLocation = "Locations";
                            txtStocking_UOM = "UOMS";
                            txtItem_description = "ItemDescription";
                            Shop_Order = "ShopOrder";
                            txtLot_First = "txtLot_First";
                            txtLot_Second = "LOT_NUMBER";
                            txtUser = "User";
                            txtQuantity = "Qtys";
                            txtReference_Number = "Refs";
                            txtItem_Type = "Itemtype";
                            txtCustNumber = "CustNumber";
                            txtBackNumber = "BackNumber";
                            DateTime_now = "ToDate";
                            string TIME = "ToTime";

                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "718-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Opened New Template");
                                documents.Replace(txtItem_Number_first, Item_Number_first, false, true);
                                documents.Replace(txtItem_Number_second, Item_Number_second, false, true);
                                documents.Replace(txtWarehouse, Warehouse, false, true);
                                documents.Replace(txtLocation, Location, false, true);
                                documents.Replace(Shop_Order, ShopOrder, false, true);
                                documents.Replace(txtLot_First, Lot_First, false, true);
                                documents.Replace(txtLot_Second, Lot_Second, false, true);
                                documents.Replace(txtStocking_UOM, Stocking_UOM, false, true);
                                documents.Replace(txtItem_description, Item_description, false, true);
                                documents.Replace(txtUser, User, false, true);
                                documents.Replace(txtQuantity, (Quantity), false, true);
                                documents.Replace(txtCustNumber, CustNumber, false, true);
                                documents.Replace(txtBackNumber, Back_Number, false, true);
                                documents.Replace(txtReference_Number, Reference_Number, false, true);
                                documents.Replace(txtItem_Type, Item_Type, false, true);
                                documents.Replace(DateTime_now, D.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(TIME, D.ToString("HH:mm:ss"), false, true);
                                documents.Save(wordTemplate);
                                documents.Close();
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "718-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Start Barcode Insert");
                                        PdfPage page = pdfDocument.Pages[0];
                                        PdfCode39ExtendedBarcode barcode = new PdfCode39ExtendedBarcode();
                                        PdfCode128Barcode code128 = new PdfCode128Barcode();
                                        try
                                        {
                                            barcode.BarHeight = 18;
                                            barcode.Size = new SizeF(200, 18);
                                            barcode.Text = Item_Number_first + Item_Number_second;
                                            barcode.TextDisplayLocation = TextLocation.None;
                                            barcode.Draw(page, new PointF(25, 125));

                                            code128 = new PdfCode128Barcode();
                                            code128.BarHeight = 16;
                                            code128.Size = new SizeF(60, 14);
                                            code128.Text = Quantity;
                                            code128.TextDisplayLocation = TextLocation.None;
                                            code128.Draw(page, new PointF(310, 158));

                                            code128 = new PdfCode128Barcode();
                                            code128.BarHeight = 16;
                                            code128.Size = new SizeF(280, 18);
                                            //barcode.NarrowBarWidth = 0.6F;
                                            code128.Text = CustNumber;
                                            code128.TextDisplayLocation = TextLocation.None;
                                            code128.Draw(page, new PointF(23, 200));
                                            if (Lot_Second != "")
                                            {


                                                barcode = new PdfCode39ExtendedBarcode();
                                                barcode.BarHeight = 16;
                                                barcode.Size = new SizeF(130, 16);
                                               //barcode.Size = new SizeF(120, 16);
                                                barcode.Text = Lot_Second;
                                                barcode.TextDisplayLocation = TextLocation.None;
                                                barcode.Draw(page, new PointF(230, 97));
                                            }
                                            barcode = new PdfCode39ExtendedBarcode();
                                            barcode.BarHeight = 16;
                                            barcode.Size = new SizeF(90, 12);
                                            barcode.Text = ShopOrder;
                                            barcode.TextDisplayLocation = TextLocation.None;
                                            barcode.Draw(page, new PointF(190, 77));


                                        }
                                        catch (Exception err)
                                        {

                                            throw;
                                        }

                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "718-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "718-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "718-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "718-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }
    }

    public class LX718RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LX718LabelReareangeSettings
    {
        public LX718RearrangeSetting NumberOfCopies
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 10,
                    RearrangeRowNumber = 4
                };
            }
        }
        public LX718RearrangeSetting Printer_IP
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 44,
                    RearrangeRowNumber = 3
                };
            }
        }

        public LX718RearrangeSetting Warehouse
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 35,
                    RearrangeRowNumber = 7
                };
            }
        }

        public LX718RearrangeSetting Location
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 8
                };
            }
        }

        public LX718RearrangeSetting Reference_Number
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 2,
                    RearrangeRowNumber = 14
                };
            }
        }


        public LX718RearrangeSetting Stocking_UOM
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 10
                };
            }
        }

        public LX718RearrangeSetting Lot_First
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 8
                };
            }
        }
        public LX718RearrangeSetting Back_Number
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 57,
                    RearrangeRowNumber = 17
                };
            }
        }
        public LX718RearrangeSetting Lot_Second
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 9
                };
            }
        }
        public LX718RearrangeSetting ShopOrder
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 40,
                    RearrangeRowNumber = 5
                };
            }
        }
        public LX718RearrangeSetting CustNumber
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 57,
                    RearrangeRowNumber = 16
                };
            }
        }
        public LX718RearrangeSetting Item_Number_first
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 2,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX718RearrangeSetting Item_Number_second
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 42,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX718RearrangeSetting Quantity
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 10,
                    RearrangeRowNumber = 13
                };
            }
        }
        public LX718RearrangeSetting Item_Type
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 15
                };
            }
        }


        public LX718RearrangeSetting User
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 12
                };
            }
        }
        public LX718RearrangeSetting Item_description
        {
            get
            {
                return new LX718RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 56,
                    RearrangeRowNumber = 11
                };
            }
        }
    }
}
