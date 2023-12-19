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
    public class LX723LabelEngine : IDisposable
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
            LX723LabelReareangeSettings LX723RearrangeSettings = new LX723LabelReareangeSettings();
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                    string Item_Number_first = "";
                    try
                    {
                        Item_Number_first = lofLines[LX723RearrangeSettings.Item_Number_first.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Item_Number_first.RearrangeColumnStart, LX723RearrangeSettings.Item_Number_first.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number_first = lofLines[LX723RearrangeSettings.Item_Number_first.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Item_Number_first.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Item_Number_second = "";
                    try
                    {
                        Item_Number_second = lofLines[LX723RearrangeSettings.Item_Number_second.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Item_Number_second.RearrangeColumnStart, LX723RearrangeSettings.Item_Number_second.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number_second = lofLines[LX723RearrangeSettings.Item_Number_second.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Item_Number_second.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Reference_Number = "";
                    try
                    {
                        Reference_Number = lofLines[LX723RearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Reference_Number.RearrangeColumnStart, LX723RearrangeSettings.Reference_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Reference_Number = lofLines[LX723RearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Reference_Number.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Warehouse_From = "";
                    try
                    {
                        Warehouse_From = lofLines[LX723RearrangeSettings.Warehouse_From.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Warehouse_From.RearrangeColumnStart, LX723RearrangeSettings.Warehouse_From.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Warehouse_From = lofLines[LX723RearrangeSettings.Warehouse_From.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Warehouse_From.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Warehouse_To = "";
                    try
                    {
                        Warehouse_To = lofLines[LX723RearrangeSettings.Warehouse_To.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Warehouse_To.RearrangeColumnStart, LX723RearrangeSettings.Warehouse_To.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Warehouse_To = lofLines[LX723RearrangeSettings.Warehouse_To.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Warehouse_To.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Location_From = "";
                    try
                    {
                        Location_From = lofLines[LX723RearrangeSettings.Location_From.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Location_From.RearrangeColumnStart, LX723RearrangeSettings.Location_From.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Location_From = lofLines[LX723RearrangeSettings.Location_From.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Location_From.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string Location_To = "";
                    try
                    {
                        Location_To = lofLines[LX723RearrangeSettings.Location_To.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Location_To.RearrangeColumnStart, LX723RearrangeSettings.Location_To.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Location_To = lofLines[LX723RearrangeSettings.Location_To.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Location_To.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }




                    string Lot_First = "";
                    try
                    {
                        Lot_First = lofLines[LX723RearrangeSettings.Lot_First.RearrangeRowNumber - 2].Substring(LX723RearrangeSettings.Lot_First.RearrangeColumnStart, LX723RearrangeSettings.Lot_First.RearrangeColumnEnd).Trim();

                    }
                    catch
                    {
                        try
                        {
                            Lot_First = lofLines[LX723RearrangeSettings.Lot_First.RearrangeRowNumber - 2].Substring(LX723RearrangeSettings.Lot_First.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Lot_Second = "";
                    try
                    {
                        Lot_Second = lofLines[LX723RearrangeSettings.Lot_Second.RearrangeRowNumber - 2].Substring(LX723RearrangeSettings.Lot_Second.RearrangeColumnStart, LX723RearrangeSettings.Lot_Second.RearrangeColumnEnd).Trim();

                    }
                    catch
                    {
                        try
                        {
                            Lot_Second = lofLines[LX723RearrangeSettings.Lot_Second.RearrangeRowNumber - 2].Substring(LX723RearrangeSettings.Lot_Second.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Stocking_UOM = "";
                    try
                    {
                        Stocking_UOM = lofLines[LX723RearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Stocking_UOM.RearrangeColumnStart, LX723RearrangeSettings.Stocking_UOM.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Stocking_UOM = lofLines[LX723RearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Stocking_UOM.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }



                    string Item_description = "";
                    try
                    {
                        Item_description = lofLines[LX723RearrangeSettings.Item_description.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Item_description.RearrangeColumnStart, LX723RearrangeSettings.Item_description.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_description = lofLines[LX723RearrangeSettings.Item_description.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Item_description.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }


                    string User = "";
                    try
                    {
                        User = lofLines[LX723RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.User.RearrangeColumnStart, LX723RearrangeSettings.User.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            User = lofLines[LX723RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.User.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string BarcodeQuantity = "";
                    string Quantity = "";
                    try
                    {
                        Quantity = lofLines[LX723RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Quantity.RearrangeColumnStart, LX723RearrangeSettings.Quantity.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Quantity = lofLines[LX723RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Quantity.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    try
                    {
                        decimal tempDecimalQuantity = Convert.ToDecimal(Quantity);
                        BarcodeQuantity = tempDecimalQuantity.ToString();
                        int tempIntQuantity = Convert.ToInt32(tempDecimalQuantity);
                        Quantity = tempIntQuantity.ToString();
                        BarcodeQuantity = Quantity;
                    }
                    catch
                    {

                    }

                    string Customer_Item_Number = "";
                    try
                    {
                        Customer_Item_Number = lofLines[LX723RearrangeSettings.Customer_Item_Number.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Customer_Item_Number.RearrangeColumnStart, LX723RearrangeSettings.Customer_Item_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Customer_Item_Number = lofLines[LX723RearrangeSettings.Customer_Item_Number.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Customer_Item_Number.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string Printer_IP = "";
                    try
                    {
                        Printer_IP = lofLines[LX723RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Printer_IP.RearrangeColumnStart, LX723RearrangeSettings.Printer_IP.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Printer_IP = lofLines[LX723RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Printer_IP.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string NumberOfCopies = "";
                    try
                    {
                        NumberOfCopies = lofLines[LX723RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.NumberOfCopies.RearrangeColumnStart, LX723RearrangeSettings.NumberOfCopies.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            NumberOfCopies = lofLines[LX723RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.NumberOfCopies.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string IXDESC = "";
                    try
                    {
                        IXDESC = lofLines[LX723RearrangeSettings.IXDESC.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.IXDESC.RearrangeColumnStart, LX723RearrangeSettings.IXDESC.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            IXDESC = lofLines[LX723RearrangeSettings.IXDESC.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.IXDESC.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\723-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
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



                            string txtItem_Number_first, txtItem_Number_second, txtJobInfo_CustomerItemNo, txtJobInfo_BackNo, txtCustomer_Item_Number, txtReference_Number, txtWarehouse_From, txtWarehouse_To, txtLocation_From, txtLocation_To, txtLot_First, txtLot_Second, txtStocking_UOM, txtItem_description, txtUser, txtQuantity;
                            DateTime D = DateTime.Now;
                            txtItem_Number_first = "Itm1";
                            txtItem_Number_second = "Itm2";
                            txtReference_Number = "Ref1";
                            txtWarehouse_From = "D1";
                            txtWarehouse_To = "ToWH";
                            txtLocation_To = "ToLoc";
                            txtLocation_From = "221221";
                            txtLot_First = "Lot1";
                            txtLot_Second = "Lot2";
                            txtStocking_UOM = "UOMS";
                            txtCustomer_Item_Number = "ItemNum";

                            txtItem_description = "ItemDescription";
                            txtUser = "User";
                            txtQuantity = "Qtys";
                            txtJobInfo_CustomerItemNo = "[JobInfo:|CustomerItemNo]";
                            txtJobInfo_BackNo = "[JobInfo:|BackNo]";



                            string DateTime_now = "ToDate";
                            string TIME = "ToTime";

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Found Template");
                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "723-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Opened New Template");

                                documents.Replace(txtItem_Number_first, Item_Number_first, false, true);
                                documents.Replace(txtItem_Number_second, Item_Number_second, false, true);
                                documents.Replace(txtReference_Number, Reference_Number, false, true);
                                documents.Replace(txtWarehouse_From, Warehouse_From, false, true);
                                documents.Replace(txtWarehouse_To, Warehouse_To, false, true);
                                documents.Replace(txtLocation_To, Location_To, false, true);
                                documents.Replace(txtLocation_From, Location_From, false, true);
                                documents.Replace(txtLot_First, Lot_First, false, true);
                                documents.Replace(txtLot_Second, Lot_First, false, true);
                                documents.Replace(txtStocking_UOM, Stocking_UOM, false, true);
                                documents.Replace(txtItem_description, Item_description, false, true);
                                documents.Replace(txtUser, User, false, true);
                                documents.Replace(txtCustomer_Item_Number, Customer_Item_Number, false, true);
                                documents.Replace(txtQuantity, Quantity, false, true);

                                documents.Replace(DateTime_now, DateTime.Now.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(TIME, DateTime.Now.ToString("HH:mm:ss"), false, true);
                                documents.Replace(txtJobInfo_CustomerItemNo, "Customer Item No:", false, true);
                                documents.Replace(txtJobInfo_BackNo, "Back No:", false, true);
                                documents.Save(wordTemplate);
                                documents.Close();
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "723-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Start Barcode Insert");
                                        PdfPage page = pdfDocument.Pages[0];
                                        PdfCode39Barcode barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 12;
                                        barcode.Size = new SizeF(50, 12);
                                        barcode.Text = Warehouse_To;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(200, 65));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 14;
                                        barcode.Size = new SizeF(80, 14);
                                        barcode.Text = Location_To;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(290, 65));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 18;
                                        barcode.Size = new SizeF(230, 18);
                                        barcode.Text = Item_Number_first + Item_Number_second;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(35, 130));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 16;
                                        barcode.Size = new SizeF(160, 16);
                                        barcode.Text = Customer_Item_Number;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(35, 245));

                                        //barcode = new PdfCode39Barcode();
                                        //barcode.BarHeight = 14;
                                        //barcode.NarrowBarWidth = 0.5F;
                                        ////barcode.Size = new SizeF(80, 14);
                                        //barcode.Text = BarcodeQuantity;
                                        //barcode.TextDisplayLocation = TextLocation.None;
                                        //barcode.Draw(page, new PointF(310, 164));

                                        List<string> lofQRcodeitems = new List<string>()
                                        {
                                            "HR434NF",
                                            "HR435NF",
                                            "HR440NF01",
                                            "HR441NF01",
                                            "HR448NF",
                                            "HR461G",
                                            "HR462J",
                                            "HR463AE",
                                            "HR464AK",
                                            "HR465AE",
                                            "HR466AK",
                                            "HR467AG",
                                            "HR468AJ",
                                            "HR468AM",
                                            "HR469AG",
                                            "HR470AL",
                                            "HR471H",
                                            "HR472H",
                                            "HR473J",
                                            "HR474K",
                                            "HR475K",
                                            "HR476L",
                                            "HR477E",
                                            "HR478AN",
                                            "HR479AM",
                                            "HR481N",
                                            "HR482BC",
                                            "HR483AT",
                                            "HR484K",
                                            "HR485P",
                                            "HR486BE",
                                            "HR487BB",
                                            "HR489Q",
                                            "HR490BG",
                                            "HR492BD",
                                            "X453NF"
                                        };

                                        if(Item_Number_first == "22" && lofQRcodeitems.Contains(Item_Number_second))
                                        {
                                            PdfQRBarcode qrCode = new PdfQRBarcode();
                                            qrCode.Size = new SizeF(70, 70);
                                            //qrCode.Text = IXDESC + "," + Quantity; LOG 70140
                                            qrCode.Text = IXDESC.Trim();
                                            qrCode.Draw(page, new PointF(325, 195));
                                        }

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 14;
                                        barcode.Size = new SizeF(50, 14);
                                        barcode.Text = Quantity;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(325, 164));

                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "723-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "723-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "723-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "723-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }
    }

    public class LX723RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LX723LabelReareangeSettings
    {
        public LX723RearrangeSetting NumberOfCopies
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 10,
                    RearrangeRowNumber = 4
                };
            }
        }
        public LX723RearrangeSetting Printer_IP
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 44,
                    RearrangeRowNumber = 3
                };
            }
        }

        public LX723RearrangeSetting Item_Number_first
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 2,
                    RearrangeRowNumber = 5
                };
            }
        }
        public LX723RearrangeSetting Item_Number_second
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 5
                };
            }
        }
        public LX723RearrangeSetting Reference_Number
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX723RearrangeSetting Warehouse_From
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 35,
                    RearrangeRowNumber = 7
                };
            }
        }
        public LX723RearrangeSetting Warehouse_To
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 31,
                    RearrangeRowNumber = 8
                };
            }
        }

        public LX723RearrangeSetting Location_From
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 9
                };
            }
        }

        public LX723RearrangeSetting Location_To
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 33,
                    RearrangeRowNumber = 10
                };
            }
        }
        public LX723RearrangeSetting Lot_First
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 36,
                    RearrangeRowNumber = 12
                };
            }
        }


        public LX723RearrangeSetting Stocking_UOM
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 12
                };
            }
        }

        public LX723RearrangeSetting Item_description
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 56,
                    RearrangeRowNumber = 13
                };
            }
        }
        public LX723RearrangeSetting User
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 14
                };
            }
        }
        public LX723RearrangeSetting Quantity
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 40,
                    RearrangeRowNumber = 15
                };
            }
        }

        // === READS 1908000326

        // ===== we using lot_second =====
        public LX723RearrangeSetting Lot_Second
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 36,
                    RearrangeRowNumber = 12
                };
            }
        }


        public LX723RearrangeSetting Item_Type
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 15
                };
            }
        }





        
        public LX723RearrangeSetting Customer_Item_Number
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 52,
                    RearrangeRowNumber = 16
                };
            }
        }

        public LX723RearrangeSetting IXDESC
        {
            get
            {
                return new LX723RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 55,
                    RearrangeRowNumber = 22
                };
            }
        }


    }
}
