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
    public class LX720LabelEngine : IDisposable
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
            LX720LabelReareangeSettings LX720RearrangeSettings = new LX720LabelReareangeSettings();
            while (lofFileData.Count > 0)
            {
                try
                {
                    string[] lofLines = File.ReadAllText(lofFileData[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);


                    string Item_Number_first = "";
                    try
                    {
                        Item_Number_first = lofLines[LX720RearrangeSettings.Item_Number_first.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Item_Number_first.RearrangeColumnStart, LX720RearrangeSettings.Item_Number_first.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number_first = lofLines[LX720RearrangeSettings.Item_Number_first.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Item_Number_first.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Item_Number_second = "";
                    try
                    {
                        Item_Number_second = lofLines[LX720RearrangeSettings.Item_Number_second.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Item_Number_second.RearrangeColumnStart, LX720RearrangeSettings.Item_Number_second.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_Number_second = lofLines[LX720RearrangeSettings.Item_Number_second.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Item_Number_second.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Reference_Number = "";
                    try
                    {
                        Reference_Number = lofLines[LX720RearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Reference_Number.RearrangeColumnStart, LX720RearrangeSettings.Reference_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Reference_Number = lofLines[LX720RearrangeSettings.Reference_Number.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Reference_Number.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Warehouse_From = "";
                    try
                    {
                        Warehouse_From = lofLines[LX720RearrangeSettings.Warehouse_From.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Warehouse_From.RearrangeColumnStart, LX720RearrangeSettings.Warehouse_From.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Warehouse_From = lofLines[LX720RearrangeSettings.Warehouse_From.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Warehouse_From.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Warehouse_To = "";
                    try
                    {
                        Warehouse_To = lofLines[LX720RearrangeSettings.Warehouse_To.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Warehouse_To.RearrangeColumnStart, LX720RearrangeSettings.Warehouse_To.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Warehouse_To = lofLines[LX720RearrangeSettings.Warehouse_To.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Warehouse_To.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string Location_From = "";
                    try
                    {
                        Location_From = lofLines[LX720RearrangeSettings.Location_From.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Location_From.RearrangeColumnStart, LX720RearrangeSettings.Location_From.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Location_From = lofLines[LX720RearrangeSettings.Location_From.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Location_From.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }
                    string Location_To = "";
                    try
                    {
                        Location_To = lofLines[LX720RearrangeSettings.Location_To.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Location_To.RearrangeColumnStart, LX720RearrangeSettings.Location_To.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Location_To = lofLines[LX720RearrangeSettings.Location_To.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Location_To.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }




                    string Lot_First = "";
                    try
                    {
                        Lot_First = lofLines[LX720RearrangeSettings.Lot_First.RearrangeRowNumber - 2].Substring(LX720RearrangeSettings.Lot_First.RearrangeColumnStart, LX720RearrangeSettings.Lot_First.RearrangeColumnEnd).Trim();

                    }
                    catch
                    {
                        try
                        {
                            Lot_First = lofLines[LX720RearrangeSettings.Lot_First.RearrangeRowNumber - 2].Substring(LX720RearrangeSettings.Lot_First.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }
                    string Stocking_UOM = "";
                    try
                    {
                        Stocking_UOM = lofLines[LX720RearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Stocking_UOM.RearrangeColumnStart, LX720RearrangeSettings.Stocking_UOM.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Stocking_UOM = lofLines[LX720RearrangeSettings.Stocking_UOM.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Stocking_UOM.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }



                    string Item_description = "";
                    try
                    {
                        Item_description = lofLines[LX720RearrangeSettings.Item_description.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Item_description.RearrangeColumnStart, LX720RearrangeSettings.Item_description.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Item_description = lofLines[LX720RearrangeSettings.Item_description.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Item_description.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }


                    string User = "";
                    try
                    {
                        User = lofLines[LX720RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.User.RearrangeColumnStart, LX720RearrangeSettings.User.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            User = lofLines[LX720RearrangeSettings.User.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.User.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }


                    string Quantity = "";
                    try
                    {
                        Quantity = lofLines[LX720RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Quantity.RearrangeColumnStart, LX720RearrangeSettings.Quantity.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Quantity = lofLines[LX720RearrangeSettings.Quantity.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Quantity.RearrangeColumnStart).TrimStart();
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

                    string Customer_Item_Number = "";
                    try
                    {
                        Customer_Item_Number = lofLines[LX720RearrangeSettings.Customer_Item_Number.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Customer_Item_Number.RearrangeColumnStart, LX720RearrangeSettings.Customer_Item_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Customer_Item_Number = lofLines[LX720RearrangeSettings.Customer_Item_Number.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Customer_Item_Number.RearrangeColumnStart).TrimStart();
                        }
                        catch
                        {

                        }
                    }

                    string Printer_IP = "";
                    try
                    {
                        Printer_IP = lofLines[LX720RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Printer_IP.RearrangeColumnStart, LX720RearrangeSettings.Printer_IP.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Printer_IP = lofLines[LX720RearrangeSettings.Printer_IP.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.Printer_IP.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    string NumberOfCopies = "";
                    try
                    {
                        NumberOfCopies = lofLines[LX720RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.NumberOfCopies.RearrangeColumnStart, LX720RearrangeSettings.NumberOfCopies.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            NumberOfCopies = lofLines[LX720RearrangeSettings.NumberOfCopies.RearrangeRowNumber - 1].Substring(LX720RearrangeSettings.NumberOfCopies.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\720-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Found Template");
                            string txtItem_Number_first, txtItem_Number_second, txtCustomer_Item_Number, txtToDate, txtToTime, txtReference_Number, txtWarehouse_From, txtWarehouse_To, txtLocation_From, txtLocation_To, txtLot_First, txtStocking_UOM, txtItem_description, txtUser, txtQuantity;
                            DateTime D = DateTime.Now;
                            txtItem_Number_first = "Item1";
                            txtItem_Number_second = "Item2";
                            txtReference_Number = "Refs";
                            txtWarehouse_From = "FrmWHs";
                            txtWarehouse_To = "ToWH";
                            txtLocation_To = "ToLoc";
                            txtLocation_From = "FrmLoc";
                            txtLot_First = "lOTnUM";
                            txtStocking_UOM = "UOMs";
                            txtItem_description = "ItemDescription";
                            txtUser = "UserS";
                            txtQuantity = "Qts";
                            txtCustomer_Item_Number = "CustItemNum";
                            txtToTime = "ToTime";
                            txtToDate = "ToDate";


                            // Quantity = Convert.ToInt32(Quantity).ToString();
                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "720-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Opened New Template");
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
                                documents.Replace(txtCustomer_Item_Number, Customer_Item_Number, false, true);
                                documents.Replace("`", " ", false, true);

                                // ======== QUANTITY DOESNT WANT TO CONVERT 
                                documents.Replace(txtQuantity, Quantity, false, true);

                                documents.Replace(txtToDate, D.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(txtToTime, D.ToString("HH:mm:ss"), false, true);
                                documents.Save(wordTemplate);
                                documents.Close();
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "720-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Start Barcode Insert");
                                        PdfPage page = pdfDocument.Pages[0];

                                        PdfCode39Barcode barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 18;
                                        barcode.Size = new SizeF(230, 18);
                                        barcode.Text = Item_Number_first + Item_Number_second;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(20, 115));

                                        ////  =================================================================
                                        ////  =================================================================
                                        PdfCode128Barcode barcodes = new PdfCode128Barcode();
                                        if (Lot_First != "")
                                        {
                                            barcodes.BarHeight = 12;
                                            barcodes.Text = Lot_First;
                                            barcodes.TextDisplayLocation = TextLocation.None;
                                            barcodes.Draw(page, new PointF(200, 86));
                                        }
                                        ////  =================================================================




                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 14;
                                        barcode.Size = new SizeF(60, 14);
                                        barcode.Text = Quantity;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(315, 148));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 14;
                                        //barcode.Size = new SizeF(290, 18);
                                        //barcode.NarrowBarWidth = 0.6F;
                                        barcode.Text = Customer_Item_Number;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(20, 190));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 12;
                                        barcode.Text = Warehouse_To;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(220, 51));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 12;
                                        barcode.Text = Location_To;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(270, 68));

                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "720-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "720-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "720-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "720-LX FAILED");
                        lofFileData.RemoveAt(0);
                    }
                    catch
                    {

                    }
                }
            }
        }


        public class LX720RearrangeSetting
        {
            public int RearrangeColumnStart { get; set; }
            public int RearrangeColumnEnd { get; set; }
            public int RearrangeRowNumber { get; set; }
        }

        public class LX720LabelReareangeSettings
        {
            public LX720RearrangeSetting NumberOfCopies
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 10,
                        RearrangeRowNumber = 4
                    };
                }
            }
            public LX720RearrangeSetting Printer_IP
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 27,
                        RearrangeColumnEnd = 44,
                        RearrangeRowNumber = 3
                    };
                }
            }

            public LX720RearrangeSetting Item_Number_first
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 2,
                        RearrangeRowNumber = 5
                    };
                }
            }
            public LX720RearrangeSetting Item_Number_second
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 27,
                        RearrangeColumnEnd = 41,
                        RearrangeRowNumber = 5
                    };
                }
            }
            public LX720RearrangeSetting Reference_Number
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 41,
                        RearrangeRowNumber = 6
                    };
                }
            }
            public LX720RearrangeSetting Warehouse_From
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 35,
                        RearrangeRowNumber = 7
                    };
                }
            }
            public LX720RearrangeSetting Warehouse_To
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 35,
                        RearrangeRowNumber = 8
                    };
                }
            }

            public LX720RearrangeSetting Location_From
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 38,
                        RearrangeRowNumber = 9
                    };
                }
            }

            public LX720RearrangeSetting Location_To
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 32,
                        RearrangeRowNumber = 10
                    };
                }
            }
            public LX720RearrangeSetting Lot_First
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 36,
                        RearrangeRowNumber = 12
                    };
                }
            }


            public LX720RearrangeSetting Stocking_UOM
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 32,
                        RearrangeRowNumber = 12
                    };
                }
            }

            public LX720RearrangeSetting Item_description
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 56,
                        RearrangeRowNumber = 13
                    };
                }
            }
            public LX720RearrangeSetting User
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 32,
                        RearrangeRowNumber = 14
                    };
                }
            }
            public LX720RearrangeSetting Quantity
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 40,
                        RearrangeRowNumber = 15
                    };
                }
            }

            // === READS 1908000326

            // ===== we using lot_second =====
            public LX720RearrangeSetting Lot_Second
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 41,
                        RearrangeRowNumber = 9
                    };
                }
            }


            public LX720RearrangeSetting Item_Type
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 41,
                        RearrangeRowNumber = 15
                    };
                }
            }






            public LX720RearrangeSetting Customer_Item_Number
            {
                get
                {
                    return new LX720RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 52,
                        RearrangeRowNumber = 16
                    };
                }
            }

        }
    }
}
