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
    public class LX717LabelEngine : IDisposable
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
            LabelReareangeSettings rearrangeSettings = new LabelReareangeSettings();
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                    LX717LabelReareangeSettings LX717RearrangeSettings = new LX717LabelReareangeSettings();
                    string Warehouse = "";
                    try
                    {
                        Warehouse = lofLines[LX717RearrangeSettings.Warehouse.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Warehouse.RearrangeColumnStart, LX717RearrangeSettings.Warehouse.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Warehouse = lofLines[LX717RearrangeSettings.Warehouse.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Warehouse.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Item_Number_first = "";
                    try
                    {
                        Item_Number_first = lofLines[LX717RearrangeSettings.Item_Number_first.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Item_Number_first.RearrangeColumnStart, LX717RearrangeSettings.Item_Number_first.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number_first = lofLines[LX717RearrangeSettings.Item_Number_first.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Item_Number_first.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Item_Number_second = "";
                    try
                    {
                        Item_Number_second = lofLines[LX717RearrangeSettings.Item_Number_second.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Item_Number_second.RearrangeColumnStart, LX717RearrangeSettings.Item_Number_second.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number_second = lofLines[LX717RearrangeSettings.Item_Number_second.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Item_Number_second.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Location = "";
                    try
                    {
                        Location = lofLines[LX717RearrangeSettings.Location.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Location.RearrangeColumnStart, LX717RearrangeSettings.Location.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Location = lofLines[LX717RearrangeSettings.Location.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Location.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string User = "";
                    try
                    {
                        User = lofLines[LX717RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.User.RearrangeColumnStart, LX717RearrangeSettings.User.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            User = lofLines[LX717RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.User.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Stocking_UOM = "";
                    try
                    {
                        Stocking_UOM = lofLines[LX717RearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Stocking_UOM.RearrangeColumnStart, LX717RearrangeSettings.Stocking_UOM.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Stocking_UOM = lofLines[LX717RearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Stocking_UOM.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Lot_First = "";
                    try
                    {
                        Lot_First = lofLines[LX717RearrangeSettings.Lot_First.RearrangeRowNumber - 2].Substring(LX717RearrangeSettings.Lot_First.RearrangeColumnStart, LX717RearrangeSettings.Lot_First.RearrangeColumnEnd).Trim();
                        try { Lot_First = Convert.ToInt32(Lot_First).ToString(); } catch { }
                    }
                    catch
                    {
                        try
                        {
                            Lot_First = lofLines[LX717RearrangeSettings.Lot_First.RearrangeRowNumber - 2].Substring(LX717RearrangeSettings.Lot_First.RearrangeColumnStart).Trim();
                            try { Lot_First = Convert.ToInt32(Lot_First).ToString(); } catch { }
                        }
                        catch
                        {

                        }
                    }

                    string Item_description = "";
                    try
                    {
                        Item_description = lofLines[LX717RearrangeSettings.Item_description.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Item_description.RearrangeColumnStart, LX717RearrangeSettings.Item_description.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_description = lofLines[LX717RearrangeSettings.Item_description.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Item_description.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Lot_Second = "";
                    try
                    {
                        Lot_Second = lofLines[LX717RearrangeSettings.Lot_Second.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Lot_Second.RearrangeColumnStart, LX717RearrangeSettings.Lot_Second.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Lot_Second = lofLines[LX717RearrangeSettings.Lot_Second.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Lot_Second.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }
                    string Item_Type = "";
                    try
                    {
                        Item_Type = lofLines[LX717RearrangeSettings.Item_Type.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Item_Type.RearrangeColumnStart, LX717RearrangeSettings.Item_Type.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Type = lofLines[LX717RearrangeSettings.Item_Type.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Item_Type.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string Quantity = "";
                    try
                    {
                        Quantity = lofLines[LX717RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Quantity.RearrangeColumnStart, LX717RearrangeSettings.Quantity.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Quantity = lofLines[LX717RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Quantity.RearrangeColumnStart).TrimStart();
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

                    string CustNumber = "";
                    try
                    {
                        CustNumber = lofLines[LX717RearrangeSettings.CustNumber.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.CustNumber.RearrangeColumnStart, LX717RearrangeSettings.CustNumber.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            CustNumber = lofLines[LX717RearrangeSettings.CustNumber.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.CustNumber.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string Back_Number = "";
                    try
                    {
                        Back_Number = lofLines[LX717RearrangeSettings.Back_Number.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Back_Number.RearrangeColumnStart, LX717RearrangeSettings.Back_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Back_Number = lofLines[LX717RearrangeSettings.Back_Number.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Back_Number.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }
                    string Reference_Number = "";
                    try
                    {
                        Reference_Number = lofLines[LX717RearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Reference_Number.RearrangeColumnStart, LX717RearrangeSettings.Reference_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Reference_Number = lofLines[LX717RearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Reference_Number.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string Printer_IP = "";
                    try
                    {
                        Printer_IP = lofLines[LX717RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Printer_IP.RearrangeColumnStart, LX717RearrangeSettings.Printer_IP.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Printer_IP = lofLines[LX717RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.Printer_IP.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string NumberOfCopies = "";
                    try
                    {
                        NumberOfCopies = lofLines[LX717RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.NumberOfCopies.RearrangeColumnStart, LX717RearrangeSettings.NumberOfCopies.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            NumberOfCopies = lofLines[LX717RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX717RearrangeSettings.NumberOfCopies.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Start New Label");
                        string originalTemplateWordDocument = "";
                        if (Warehouse == "U1" || Warehouse == "U9" || Warehouse == "U2" || Warehouse == "U8")
                        {
                            originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\717A-LX Template.docx";
                        }
                        else
                        {
                            originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\717B-LX Template.docx";
                        }

                        if (File.Exists(originalTemplateWordDocument))
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Found Template");
                            string txtWarehouse, txtReference_Number, txtLocation, txtUser, txtItem_Number_first, txtItem_Number_second, txtItem_Type, txtStocking_UOM, txtItem_description, txtBackNumber, txtQuantity, txtCustNumber, txtLot_Second, txtLot_First, DateTime_now;
                            DateTime D = DateTime.Now;
                            txtItem_Number_first = "Tem";
                            txtItem_Number_second = "Item2";
                            txtWarehouse = "Whs1";
                            txtLocation = "Locations";
                            txtStocking_UOM = "UOMS";
                            txtItem_description = "ItemDescription";
                            txtLot_First = "txtLot_First";
                            txtLot_Second = "txtLot_First";
                            txtUser = "User";
                            txtQuantity = "Qtys";
                            txtReference_Number = "Refs";
                            txtItem_Type = "Itemtype";
                            txtCustNumber = "CustNumber";
                            txtBackNumber = "BackNumber";
                            DateTime_now = "ToDate";
                            string TIME = "ToTime";

                            string wordTemplate = "";
                            if (Warehouse == "U1" || Warehouse == "U9" || Warehouse == "U2" || Warehouse == "U8")
                            {
                                wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "717A-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            }
                            else
                            {
                                wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "717B-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            }

                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Opened New Template");
                                documents.Replace(txtItem_Number_first, Item_Number_first, false, true);
                                documents.Replace(txtItem_Number_second, Item_Number_second, false, true);
                                documents.Replace(txtWarehouse, Warehouse, false, true);
                                documents.Replace(txtLocation, Location, false, true);
                                documents.Replace(txtLot_First, Lot_First, false, true);
                                documents.Replace(txtLot_Second, Lot_Second, false, true);
                                //documents.Replace(DateTime_now, D.ToShortDateString(), false, true);
                                documents.Replace(txtStocking_UOM, Stocking_UOM, false, true);
                                documents.Replace(txtItem_description, Item_description, false, true);
                                documents.Replace(txtUser, User, false, true);
                                documents.Replace(txtQuantity, Quantity, false, true);
                                documents.Replace(txtCustNumber, CustNumber, false, true);
                                documents.Replace(txtBackNumber, Back_Number, false, true);
                                documents.Replace(txtReference_Number, Reference_Number, false, true);
                                documents.Replace(txtItem_Type, Item_Type, false, true);
                                documents.Replace(DateTime_now, D.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(TIME, D.ToString("HH:mm:ss"), false, true);

                                documents.Replace("BARCODE1", Item_Number_first + Item_Number_second, false, true);

                                documents.Save(wordTemplate);
                                documents.Close();
                            }


                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Saved and Closed Template");

                            string newPDFFileName = "";
                            if (Warehouse == "U1" || Warehouse == "U9" || Warehouse == "U2" || Warehouse == "U8")
                            {
                                newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "717A-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            }
                            else
                            {
                                newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "717B-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Converted and Saved PDF Document");


                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Start Barcode Insert");
                                        PdfPage page = pdfDocument.Pages[0];
                                        if (Warehouse == "U1" || Warehouse == "U9" || Warehouse == "U2" || Warehouse == "U8")
                                        {
                                            PdfCode39Barcode barcode = new PdfCode39Barcode();
                                            barcode.BarHeight = 20;
                                            barcode.Size = new SizeF(200, 18);
                                            barcode.Text = Item_Number_first + Item_Number_second;
                                            barcode.TextDisplayLocation = TextLocation.None;
                                            barcode.Draw(page, new PointF(30, 165));

                                            barcode = new PdfCode39Barcode();
                                            barcode.BarHeight = 14;
                                            barcode.Size = new SizeF(60, 14);
                                            barcode.Text = Quantity;
                                            barcode.TextDisplayLocation = TextLocation.None;
                                            barcode.Draw(page, new PointF(230, 198));

                                            barcode = new PdfCode39Barcode();
                                            barcode.BarHeight = 16;
                                            barcode.Size = new SizeF(250, 18);
                                            //  barcode.NarrowBarWidth = 0.6F;
                                            barcode.Text = CustNumber;
                                            barcode.TextDisplayLocation = TextLocation.None;
                                            barcode.Draw(page, new PointF(30, 250));
                                        }
                                        else
                                        {
                                            PdfCode39Barcode barcode = new PdfCode39Barcode();
                                            barcode.BarHeight = 20;
                                            barcode.Size = new SizeF(240, 18);
                                            barcode.Text = Item_Number_first + Item_Number_second;
                                            barcode.TextDisplayLocation = TextLocation.None;
                                            barcode.Draw(page, new PointF(25, 125));

                                            barcode = new PdfCode39Barcode();
                                            barcode.BarHeight = 14;
                                            barcode.Size = new SizeF(60, 14);
                                            barcode.Text = Quantity;
                                            barcode.TextDisplayLocation = TextLocation.None;
                                            barcode.Draw(page, new PointF(315, 158));

                                            barcode = new PdfCode39Barcode();
                                            barcode.BarHeight = 16;
                                            barcode.Size = new SizeF(290, 18);
                                            //  barcode.NarrowBarWidth = 0.6F;
                                            barcode.Text = CustNumber;
                                            barcode.TextDisplayLocation = TextLocation.None;
                                            barcode.Draw(page, new PointF(23, 200));
                                        }

                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = "";
                            if (Warehouse == "U1" || Warehouse == "U9" || Warehouse == "U2" || Warehouse == "U8")
                            {
                                outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "717-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
                            }
                            else
                            {
                                outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "717-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
                            }
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "717-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "717-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "717-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }
    }

    public class LX717RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LX717LabelReareangeSettings
    {
        public LX717RearrangeSetting NumberOfCopies
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 10,
                    RearrangeRowNumber = 4
                };
            }
        }
        public LX717RearrangeSetting Printer_IP
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 44,
                    RearrangeRowNumber = 3
                };
            }
        }

        public LX717RearrangeSetting Warehouse
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 35,
                    RearrangeRowNumber = 7
                };
            }
        }

        public LX717RearrangeSetting Location
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 8
                };
            }
        }

        public LX717RearrangeSetting Reference_Number
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 2,
                    RearrangeRowNumber = 14
                };
            }
        }


        public LX717RearrangeSetting Stocking_UOM
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 10
                };
            }
        }

        public LX717RearrangeSetting Lot_First
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 35,
                    RearrangeRowNumber = 9
                };
            }
        }
        public LX717RearrangeSetting Back_Number
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 40,
                    RearrangeRowNumber = 17
                };
            }
        }
        public LX717RearrangeSetting Lot_Second
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 9
                };
            }
        }
        //public LX717RearrangeSetting Time_Now
        //{
        //    get
        //    {
        //        return new LX717RearrangeSetting()
        //        {
        //            RearrangeColumnStart = 27,
        //            RearrangeColumnEnd = 5,
        //            RearrangeRowNumber = 22
        //        };
        //    }
        //}
        public LX717RearrangeSetting CustNumber
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 50,
                    RearrangeRowNumber = 16
                };
            }
        }
        public LX717RearrangeSetting Item_Number_first
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 2,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX717RearrangeSetting Item_Number_second
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 42,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX717RearrangeSetting Quantity
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 10,
                    RearrangeRowNumber = 13
                };
            }
        }
        public LX717RearrangeSetting Item_Type
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 15
                };
            }
        }


        public LX717RearrangeSetting User
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 12
                };
            }
        }
        public LX717RearrangeSetting Item_description
        {
            get
            {
                return new LX717RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 56,
                    RearrangeRowNumber = 11
                };
            }
        }
    }
}
