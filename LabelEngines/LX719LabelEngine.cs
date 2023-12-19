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
    public class LX719LabelEngine : IDisposable
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
            LX719LabelReareangeSettings LX719RearrangeSettings = new LX719LabelReareangeSettings();
            LX719LabelReareangeSettingsB LX719RearrangeSettingsB = new LX719LabelReareangeSettingsB();
            LogEngine logEngine = new LogEngine();
            FileEngine csFileInputEngine = new FileEngine();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                    string Quantity = "";
                    string NumberOfCopies = "";
                    string Item_Number_first = "";
                    string Item_Number_second = "";
                    string Reference_Number = "";
                    string Warehouse_From = "";
                    string Warehouse_To = "";
                    string Location_To = "";
                    string Location_From = "";
                    string Lot_First = "";
                    string Stocking_UOM = "";
                    string Item_description = "";
                    string User = "";
                    string Printer_IP = "";
                    if (lofLines[1].Substring(25) == "Barcode Scanning 2023")
                    {
                        try
                        {
                            Item_Number_first = lofLines[LX719RearrangeSettingsB.Item_Number_first.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Item_Number_first.RearrangeColumnStart, LX719RearrangeSettingsB.Item_Number_first.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Item_Number_first = lofLines[LX719RearrangeSettingsB.Item_Number_first.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Item_Number_first.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Item_Number_second = lofLines[LX719RearrangeSettingsB.Item_Number_second.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Item_Number_second.RearrangeColumnStart, LX719RearrangeSettingsB.Item_Number_second.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Item_Number_second = lofLines[LX719RearrangeSettingsB.Item_Number_second.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Item_Number_second.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Reference_Number = lofLines[LX719RearrangeSettingsB.Reference_Number.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Reference_Number.RearrangeColumnStart, LX719RearrangeSettingsB.Reference_Number.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Reference_Number = lofLines[LX719RearrangeSettingsB.Reference_Number.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Reference_Number.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Warehouse_From = lofLines[LX719RearrangeSettingsB.Warehouse_From.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Warehouse_From.RearrangeColumnStart, LX719RearrangeSettingsB.Warehouse_From.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Warehouse_From = lofLines[LX719RearrangeSettingsB.Warehouse_From.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Warehouse_From.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Warehouse_To = lofLines[LX719RearrangeSettingsB.Warehouse_To.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Warehouse_To.RearrangeColumnStart, LX719RearrangeSettingsB.Warehouse_To.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Warehouse_To = lofLines[LX719RearrangeSettingsB.Warehouse_To.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Warehouse_To.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Location_From = lofLines[LX719RearrangeSettingsB.Location_From.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Location_From.RearrangeColumnStart, LX719RearrangeSettingsB.Location_From.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Location_From = lofLines[LX719RearrangeSettingsB.Location_From.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Location_From.RearrangeColumnStart).TrimStart();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Location_To = lofLines[LX719RearrangeSettingsB.Location_To.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Location_To.RearrangeColumnStart, LX719RearrangeSettingsB.Location_To.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Location_To = lofLines[LX719RearrangeSettingsB.Location_To.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Location_To.RearrangeColumnStart).TrimStart();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Lot_First = lofLines[LX719RearrangeSettingsB.Lot_First.RearrangeRowNumber - 2].Substring(LX719RearrangeSettingsB.Lot_First.RearrangeColumnStart, LX719RearrangeSettingsB.Lot_First.RearrangeColumnEnd).Trim();
                            try { Lot_First = Convert.ToInt32(Lot_First).ToString(); } catch { }
                        }
                        catch
                        {
                            try
                            {
                                Lot_First = lofLines[LX719RearrangeSettingsB.Lot_First.RearrangeRowNumber - 2].Substring(LX719RearrangeSettingsB.Lot_First.RearrangeColumnStart).Trim();
                                try { Lot_First = Convert.ToInt32(Lot_First).ToString(); } catch { }
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Stocking_UOM = lofLines[LX719RearrangeSettingsB.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Stocking_UOM.RearrangeColumnStart, LX719RearrangeSettingsB.Stocking_UOM.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Stocking_UOM = lofLines[LX719RearrangeSettingsB.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Stocking_UOM.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Item_description = lofLines[LX719RearrangeSettingsB.Item_description.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Item_description.RearrangeColumnStart, LX719RearrangeSettingsB.Item_description.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Item_description = lofLines[LX719RearrangeSettingsB.Item_description.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Item_description.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            User = lofLines[LX719RearrangeSettingsB.User.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.User.RearrangeColumnStart, LX719RearrangeSettingsB.User.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                User = lofLines[LX719RearrangeSettingsB.User.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.User.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        
                        try
                        {
                            Quantity = lofLines[LX719RearrangeSettingsB.Quantity.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Quantity.RearrangeColumnStart, LX719RearrangeSettingsB.Quantity.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Quantity = lofLines[LX719RearrangeSettingsB.Quantity.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Quantity.RearrangeColumnStart).TrimStart();
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
                        try
                        {
                            Printer_IP = lofLines[LX719RearrangeSettingsB.Printer_IP.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Printer_IP.RearrangeColumnStart, LX719RearrangeSettingsB.Printer_IP.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {
                            try
                            {
                                Printer_IP = lofLines[LX719RearrangeSettingsB.Printer_IP.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.Printer_IP.RearrangeColumnStart).Trim();
                            }
                            catch
                            {

                            }
                        }
                        try
                        {
                            NumberOfCopies = lofLines[LX719RearrangeSettingsB.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.NumberOfCopies.RearrangeColumnStart, LX719RearrangeSettingsB.NumberOfCopies.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {
                            try
                            {
                                NumberOfCopies = lofLines[LX719RearrangeSettingsB.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX719RearrangeSettingsB.NumberOfCopies.RearrangeColumnStart).Trim();
                            }
                            catch
                            {

                            }
                        }
                    }
                    else
                    {
                        try
                        {
                            Item_Number_first = lofLines[LX719RearrangeSettings.Item_Number_first.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Item_Number_first.RearrangeColumnStart, LX719RearrangeSettings.Item_Number_first.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Item_Number_first = lofLines[LX719RearrangeSettings.Item_Number_first.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Item_Number_first.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Item_Number_second = lofLines[LX719RearrangeSettings.Item_Number_second.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Item_Number_second.RearrangeColumnStart, LX719RearrangeSettings.Item_Number_second.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Item_Number_second = lofLines[LX719RearrangeSettings.Item_Number_second.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Item_Number_second.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }

                        try
                        {
                            Reference_Number = lofLines[LX719RearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Reference_Number.RearrangeColumnStart, LX719RearrangeSettings.Reference_Number.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Reference_Number = lofLines[LX719RearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Reference_Number.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }

                        try
                        {
                            Warehouse_From = lofLines[LX719RearrangeSettings.Warehouse_From.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Warehouse_From.RearrangeColumnStart, LX719RearrangeSettings.Warehouse_From.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Warehouse_From = lofLines[LX719RearrangeSettings.Warehouse_From.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Warehouse_From.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Warehouse_To = lofLines[LX719RearrangeSettings.Warehouse_To.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Warehouse_To.RearrangeColumnStart, LX719RearrangeSettings.Warehouse_To.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Warehouse_To = lofLines[LX719RearrangeSettings.Warehouse_To.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Warehouse_To.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Location_From = lofLines[LX719RearrangeSettings.Location_From.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Location_From.RearrangeColumnStart, LX719RearrangeSettings.Location_From.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Location_From = lofLines[LX719RearrangeSettings.Location_From.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Location_From.RearrangeColumnStart).TrimStart();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Location_To = lofLines[LX719RearrangeSettings.Location_To.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Location_To.RearrangeColumnStart, LX719RearrangeSettings.Location_To.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Location_To = lofLines[LX719RearrangeSettings.Location_To.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Location_To.RearrangeColumnStart).TrimStart();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Lot_First = lofLines[LX719RearrangeSettings.Lot_First.RearrangeRowNumber - 2].Substring(LX719RearrangeSettings.Lot_First.RearrangeColumnStart, LX719RearrangeSettings.Lot_First.RearrangeColumnEnd).Trim();
                            try { Lot_First = Convert.ToInt32(Lot_First).ToString(); } catch { }
                        }
                        catch
                        {
                            try
                            {
                                Lot_First = lofLines[LX719RearrangeSettings.Lot_First.RearrangeRowNumber - 2].Substring(LX719RearrangeSettings.Lot_First.RearrangeColumnStart).Trim();
                                try { Lot_First = Convert.ToInt32(Lot_First).ToString(); } catch { }
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Stocking_UOM = lofLines[LX719RearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Stocking_UOM.RearrangeColumnStart, LX719RearrangeSettings.Stocking_UOM.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Stocking_UOM = lofLines[LX719RearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Stocking_UOM.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }


                        try
                        {
                            Item_description = lofLines[LX719RearrangeSettings.Item_description.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Item_description.RearrangeColumnStart, LX719RearrangeSettings.Item_description.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Item_description = lofLines[LX719RearrangeSettings.Item_description.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Item_description.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }

                        try
                        {
                            User = lofLines[LX719RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.User.RearrangeColumnStart, LX719RearrangeSettings.User.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                User = lofLines[LX719RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.User.RearrangeColumnStart).Trim();
                            }
                            catch
                            {


                            }
                        }
                        try
                        {
                            Quantity = lofLines[LX719RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Quantity.RearrangeColumnStart, LX719RearrangeSettings.Quantity.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {

                            try
                            {
                                Quantity = lofLines[LX719RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Quantity.RearrangeColumnStart).TrimStart();
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
                        try
                        {
                            Printer_IP = lofLines[LX719RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Printer_IP.RearrangeColumnStart, LX719RearrangeSettings.Printer_IP.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {
                            try
                            {
                                Printer_IP = lofLines[LX719RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.Printer_IP.RearrangeColumnStart).Trim();
                            }
                            catch
                            {

                            }
                        }

                        try
                        {
                            NumberOfCopies = lofLines[LX719RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.NumberOfCopies.RearrangeColumnStart, LX719RearrangeSettings.NumberOfCopies.RearrangeColumnEnd).Trim();
                        }
                        catch
                        {
                            try
                            {
                                NumberOfCopies = lofLines[LX719RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX719RearrangeSettings.NumberOfCopies.RearrangeColumnStart).Trim();
                            }
                            catch
                            {

                            }
                        }
                    }
                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\719-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
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

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Found Template");
                            string txtItem_Number_first, txtItem_Number_second, txtJobInfo_CustomerItemNo, txtJobInfo_BackNo, txtReference_Number, txtWarehouse_From, txtWarehouse_To, txtLocation_From, txtLocation_To, txtLot_First, txtStocking_UOM, txtItem_description, txtUser, txtQuantity;
                            DateTime D = DateTime.Now;
                            txtItem_Number_first = "Tem";
                            txtItem_Number_second = "Item2";
                            txtReference_Number = "Refs";
                            txtWarehouse_From = "Whs1";
                            txtWarehouse_To = "ToWHs";
                            txtLocation_To = "toLoc";
                            txtLocation_From = "Locations";
                            txtLot_First = "LOT_NUMBER";
                            txtStocking_UOM = "UOMs";
                            txtItem_description = "ItemDescription";
                            txtUser = "User";
                            txtQuantity = "Qtys";
                            txtJobInfo_CustomerItemNo = "[JobInfo:|CustomerItemNo]";
                            txtJobInfo_BackNo = "[JobInfo:|BackNo]";


                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "719-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Opened New Template");
                                documents.Replace(txtItem_Number_first, Item_Number_first, false, true);
                                documents.Replace(txtItem_Number_second, Item_Number_second, false, true);
                                documents.Replace(txtReference_Number, Reference_Number, false, true);
                                documents.Replace(txtWarehouse_From, Warehouse_From, false, true);
                                documents.Replace(txtWarehouse_To, Warehouse_To, false, true);
                                documents.Replace(txtLocation_To, Location_To, false, true);
                                documents.Replace(txtLocation_From, Location_From, false, true);
                                documents.Replace(txtLot_First, Lot_First, false, true);
                                documents.Replace(txtStocking_UOM, Stocking_UOM, false, true);
                                documents.Replace(txtItem_description, Item_description, false, true);
                                documents.Replace(txtUser, User, false, true);

                                // ======== QUANTITY DOESNT WANT TO CONVERT 
                                documents.Replace(txtQuantity, Quantity, false, true);

                                documents.Replace("ToDate", D.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace("ToTime", D.ToString("HH:mm:ss"), false, true);
                                documents.Replace(txtJobInfo_CustomerItemNo, "1", false, true);
                                documents.Replace(txtJobInfo_BackNo, "..", false, true);
                                documents.Save(wordTemplate);
                                documents.Close();
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "719-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Start Barcode Insert");
                                        PdfPage page = pdfDocument.Pages[0];
                                        PdfCode39Barcode barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 18;
                                        barcode.Size = new SizeF(200, 18);
                                        barcode.Text = Item_Number_first + Item_Number_second;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(25, 125));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 14;
                                        barcode.Size = new SizeF(60, 14);
                                        barcode.Text = Quantity;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(310, 158));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 16;
                                        barcode.Size = new SizeF(200, 16);
                                        barcode.Text = Lot_First;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(165, 97));

                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "719-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "719-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "719-LX");
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
    }

    public class LX719RearrangeSetting
    {
        public int RearrangeColumnStart { get; set; }
        public int RearrangeColumnEnd { get; set; }
        public int RearrangeRowNumber { get; set; }
    }

    public class LX719LabelReareangeSettings
    {
        public LX719RearrangeSetting NumberOfCopies
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 10,
                    RearrangeRowNumber = 4
                };
            }

        }
        public LX719RearrangeSetting Printer_IP
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 44,
                    RearrangeRowNumber = 3
                };
            }
        }

        public LX719RearrangeSetting Item_Number_first
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 2,
                    RearrangeRowNumber = 5
                };
            }
        }
        public LX719RearrangeSetting Item_Number_second
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 5
                };
            }
        }
        public LX719RearrangeSetting Reference_Number
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX719RearrangeSetting Warehouse_From
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 35,
                    RearrangeRowNumber = 7
                };
            }
        }
        public LX719RearrangeSetting Warehouse_To
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 27,
                    RearrangeRowNumber = 8
                };
            }
        }

        public LX719RearrangeSetting Location_From
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 9
                };
            }
        }

        public LX719RearrangeSetting Location_To
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 10
                };
            }
        }
        public LX719RearrangeSetting Lot_First
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 36,
                    RearrangeRowNumber = 12
                };
            }
        }


        public LX719RearrangeSetting Stocking_UOM
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 12
                };
            }
        }

        public LX719RearrangeSetting Item_description
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 56,
                    RearrangeRowNumber = 13
                };
            }
        }
        public LX719RearrangeSetting User
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 14
                };
            }
        }
        public LX719RearrangeSetting Quantity
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 40,
                    RearrangeRowNumber = 15
                };
            }
        }

        // === READS 1908000326

        // ===== we using lot_second =====
        public LX719RearrangeSetting Lot_Second
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 9
                };
            }
        }


        public LX719RearrangeSetting Item_Type
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 15
                };
            }
        }






        public LX719RearrangeSetting Customer_Item_Number
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 26,
                    RearrangeColumnEnd = 52,
                    RearrangeRowNumber = 16
                };
            }
        }



    }


    public class LX719LabelReareangeSettingsB
    {
        public LX719RearrangeSetting NumberOfCopies
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 10,
                    RearrangeRowNumber = 5
                };
            }

        }
        public LX719RearrangeSetting Printer_IP
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 44,
                    RearrangeRowNumber = 4
                };
            }
        }

        public LX719RearrangeSetting Item_Number_first
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 2,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX719RearrangeSetting Item_Number_second
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 27,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 6
                };
            }
        }
        public LX719RearrangeSetting Reference_Number
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 7
                };
            }
        }
        public LX719RearrangeSetting Warehouse_From
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 35,
                    RearrangeRowNumber = 8
                };
            }
        }
        public LX719RearrangeSetting Warehouse_To
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 27,
                    RearrangeRowNumber = 9
                };
            }
        }

        public LX719RearrangeSetting Location_From
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 38,
                    RearrangeRowNumber = 10
                };
            }
        }

        public LX719RearrangeSetting Location_To
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 11
                };
            }
        }
        public LX719RearrangeSetting Lot_First
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 36,
                    RearrangeRowNumber = 12
                };
            }
        }


        public LX719RearrangeSetting Stocking_UOM
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 13
                };
            }
        }

        public LX719RearrangeSetting Item_description
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 56,
                    RearrangeRowNumber = 14
                };
            }
        }
        public LX719RearrangeSetting User
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 32,
                    RearrangeRowNumber = 15
                };
            }
        }
        public LX719RearrangeSetting Quantity
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 40,
                    RearrangeRowNumber = 16
                };
            }
        }

        // === READS 1908000326

        // ===== we using lot_second =====
        public LX719RearrangeSetting Lot_Second
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 9
                };
            }
        }


        public LX719RearrangeSetting Item_Type
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 25,
                    RearrangeColumnEnd = 41,
                    RearrangeRowNumber = 15
                };
            }
        }






        public LX719RearrangeSetting Customer_Item_Number
        {
            get
            {
                return new LX719RearrangeSetting()
                {
                    RearrangeColumnStart = 26,
                    RearrangeColumnEnd = 52,
                    RearrangeRowNumber = 16
                };
            }
        }



    }
}
