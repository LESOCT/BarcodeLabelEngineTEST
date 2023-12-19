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
    public class LX702LabelEngine : IDisposable
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
            LabelReareangeSettings rearrangeSettings = new LabelReareangeSettings();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                    LX702LabelReareangeSettings LX723RearrangeSettings = new LX702LabelReareangeSettings();

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

                    string Lot_Number = "";
                    try
                    {
                        Lot_Number = lofLines[LX723RearrangeSettings.Lot_Number.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Lot_Number.RearrangeColumnStart, LX723RearrangeSettings.Lot_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Lot_Number = lofLines[LX723RearrangeSettings.Lot_Number.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Lot_Number.RearrangeColumnStart).TrimStart();
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
                        Lot_First = lofLines[LX723RearrangeSettings.Lot_First.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Lot_First.RearrangeColumnStart, LX723RearrangeSettings.Lot_First.RearrangeColumnEnd).Trim();
                        //    try { Lot_First = Convert.ToInt32(Lot_First).ToString(); } catch { }
                    }
                    catch
                    {
                        try
                        {
                            Lot_First = lofLines[LX723RearrangeSettings.Lot_First.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Lot_First.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Lot_Second = "";
                    try
                    {
                        Lot_Second = lofLines[LX723RearrangeSettings.Lot_Second.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Lot_Second.RearrangeColumnStart, LX723RearrangeSettings.Lot_Second.RearrangeColumnEnd).Trim();

                    }
                    catch
                    {
                        try
                        {
                            Lot_First = lofLines[LX723RearrangeSettings.Lot_Second.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Lot_Second.RearrangeColumnStart).Trim();
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
                    string Item_Type = "";
                    try
                    {
                        Item_Type = lofLines[LX723RearrangeSettings.Item_Type.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Item_Type.RearrangeColumnStart, LX723RearrangeSettings.Item_Type.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Type = lofLines[LX723RearrangeSettings.Item_Type.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Item_Type.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }
                    string Standard_Qty = "";
                    try
                    {
                        Standard_Qty = lofLines[LX723RearrangeSettings.Standard_Qty.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Standard_Qty.RearrangeColumnStart, LX723RearrangeSettings.Standard_Qty.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Standard_Qty = lofLines[LX723RearrangeSettings.Standard_Qty.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Standard_Qty.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }
                    string IXDESC2 = "";
                    try
                    {
                        IXDESC2 = lofLines[LX723RearrangeSettings.IXDESC2.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.IXDESC2.RearrangeColumnStart, LX723RearrangeSettings.IXDESC2.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            IXDESC2 = lofLines[LX723RearrangeSettings.IXDESC2.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.IXDESC2.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string Barcode1 = "";
                    try
                    {
                        Barcode1 = lofLines[LX723RearrangeSettings.Barcode1.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Barcode1.RearrangeColumnStart, LX723RearrangeSettings.Barcode1.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Barcode1 = lofLines[LX723RearrangeSettings.Barcode1.RearrangeRowNumber - 1].Substring(LX723RearrangeSettings.Barcode1.RearrangeColumnStart).TrimStart();
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

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\702-LX Template.docx";
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

                            string txtItem_Number_first, txtItem_Number_second, txtIXDES2, txtStandard_Qty, txtItemType, txtCustomer_Item_Number, txtWarehouse_From, txtWarehouse_To, txtLot_Number, txtLocation_To, txtItem_description, txtUser, txtQuantity;
                            DateTime D = DateTime.Now;
                            txtItem_Number_first = "It1";
                            txtItem_Number_second = "It2";
                            txtLocation_To = "Location";
                            txtLot_Number = "LotNum";
                            txtWarehouse_From = "FrmWhS";
                            txtWarehouse_To = "ToWH";
                            txtItem_description = "ItemDescription";
                            txtCustomer_Item_Number = "CustNum";
                            txtUser = "User";
                            txtItemType = "iTy";
                            txtIXDES2 = "IXDES2";
                            txtQuantity = "Qtys";
                            txtStandard_Qty = "StdQty";

                            string DateTime_now = "ToDate";
                            string TIME = "ToTime";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Found Template");
                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "702-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Opened New Template");

                                documents.Replace(txtItem_Number_first, Item_Number_first, false, true);
                                documents.Replace(txtItem_Number_second, Item_Number_second, false, true);
                                documents.Replace(txtLocation_To, Location_To, false, true);
                                documents.Replace(txtLot_Number, Lot_Number, false, true);
                                documents.Replace(txtWarehouse_From, Warehouse_From, false, true);
                                documents.Replace(txtWarehouse_To, Warehouse_To, false, true);
                                documents.Replace(txtItem_description, Item_description, false, true);
                                documents.Replace(txtCustomer_Item_Number, Customer_Item_Number, false, true);
                                documents.Replace(txtItemType, Item_Type, false, true);
                                documents.Replace(txtUser, User, false, true);
                                documents.Replace(txtIXDES2, IXDESC2, false, true);
                                documents.Replace(txtQuantity, Quantity, false, true);
                                documents.Replace(DateTime_now, D.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(TIME, D.ToString("HH:mm:ss"), false, true);
                                documents.Replace(txtStandard_Qty, Standard_Qty, false, true);

                                // ======== QUANTITY DOESNT WANT TO CONVERT 



                                documents.Save(wordTemplate);
                                documents.Close();
                            }



                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "702-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Start Barcode Insert");
                                        PdfPage page = pdfDocument.Pages[0];
                                        PdfCode39Barcode barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 12;
                                        barcode.Text = Item_Number_first + Item_Number_second;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(24, 121));

                                        PdfCode39Barcode barcode1 = new PdfCode39Barcode();
                                        barcode1.BarHeight = 12;
                                        barcode1.Text = Customer_Item_Number;
                                        barcode1.TextDisplayLocation = TextLocation.None;
                                        barcode1.Draw(page, new PointF(35, 195));

                                        PdfCode39Barcode barcode4 = new PdfCode39Barcode();
                                        barcode4.BarHeight = 15;
                                        barcode4.Text = Quantity;
                                        barcode4.TextDisplayLocation = TextLocation.None;
                                        barcode4.Draw(page, new PointF(225, 150));

                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "702-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "702-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "702-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "702-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }
    }

    public class LX702RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LX702LabelReareangeSettings
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

        public LX702RearrangeSetting Barcode1
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 43,
                    RearrangeRowNumber = 6
                };
            }

        }
        public LX702RearrangeSetting Item_Number_first
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 2,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX702RearrangeSetting Item_Number_second
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 29,
                    RearrangeColumnEnd = 14,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX702RearrangeSetting Std_Container_Qty
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 44,
                    RearrangeRowNumber = 26
                };
            }
        }
        public LX702RearrangeSetting Warehouse_From
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 35,
                    RearrangeRowNumber = 7
                };
            }
        }
        public LX702RearrangeSetting Warehouse_To
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 35,
                    RearrangeRowNumber = 10
                };
            }
        }
        public LX702RearrangeSetting Item_Type
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 1,
                    RearrangeRowNumber = 27
                };
            }
        }

        public LX702RearrangeSetting Lot_Number
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 9
                };
            }
        }

        public LX702RearrangeSetting Location_To
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 6,
                    RearrangeRowNumber = 8
                };
            }
        }
        public LX702RearrangeSetting Lot_First
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 36,
                    RearrangeRowNumber = 12
                };
            }
        }


        public LX702RearrangeSetting Stocking_UOM
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 12
                };
            }
        }

        public LX702RearrangeSetting Item_description
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 55,
                    RearrangeRowNumber = 11
                };
            }
        }
        public LX702RearrangeSetting User
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 37,
                    RearrangeRowNumber = 19
                };
            }
        }

        public LX702RearrangeSetting Quantity
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 20,
                    RearrangeRowNumber = 20
                };
            }
        }

        // === READS 1908000326

        // ===== we using lot_second =====
        public LX702RearrangeSetting Lot_Second
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 36,
                    RearrangeRowNumber = 12
                };
            }
        }




        public LX702RearrangeSetting Standard_Qty
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 44,
                    RearrangeRowNumber = 26
                };
            }
        }



        public LX702RearrangeSetting Customer_Item_Number
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 39,
                    RearrangeRowNumber = 12
                };
            }
        }
        public LX702RearrangeSetting IXDESC2
        {
            get
            {
                return new LX702RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 61,
                    RearrangeRowNumber = 28
                };
            }
        }
    }
}
