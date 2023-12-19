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
    public class LX721LabelEngine : IDisposable
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
            LabelReareangeSettings rearrangeSettings = new LabelReareangeSettings();
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

                    string Customer_Number = "";
                    try
                    {
                        Customer_Number = lofLines[rearrangeSettings.Customer_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.Customer_Number.RearrangeColumnStart, rearrangeSettings.Customer_Number.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Customer_Number = lofLines[rearrangeSettings.Customer_Number.RearrangeRowNumber - 1].Substring(rearrangeSettings.Customer_Number.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }


                    string Customer_Name = "";
                    try
                    {
                        Customer_Name = lofLines[rearrangeSettings.Customer_Name.RearrangeRowNumber - 1].Substring(rearrangeSettings.Customer_Name.RearrangeColumnStart, rearrangeSettings.Customer_Name.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            Customer_Name = lofLines[rearrangeSettings.Customer_Name.RearrangeRowNumber - 1].Substring(rearrangeSettings.Customer_Name.RearrangeColumnStart).Trim();
                        }
                        catch
                        {

                        }
                    }

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

                    string customerItemNumber = "";
                    try
                    {
                        customerItemNumber = lofLines[rearrangeSettings.CustomerItemNumber.RearrangeRowNumber - 1].Substring(rearrangeSettings.CustomerItemNumber.RearrangeColumnStart, rearrangeSettings.CustomerItemNumber.RearrangeColumnEnd).Trim();
                    }
                    catch
                    {
                        try
                        {
                            customerItemNumber = lofLines[rearrangeSettings.CustomerItemNumber.RearrangeRowNumber - 1].Substring(rearrangeSettings.CustomerItemNumber.RearrangeColumnStart).Trim();
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Start New Label");
                        string originalTemplateWordDocument = ConfigurationManager.AppSettings["LabelTemplateFolder"] + @"\721-LX Template.docx";
                        if (File.Exists(originalTemplateWordDocument))
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Found Template");

                            string txtReference_Number, txtStocking_UOM, txtItem_Description, txtQuantity, txtCustomer_Number, txtCustomer_Name, txtPO_Number, txtItem_Number, txtDate, txtTime, custItemNo;
                            DateTime datetime = DateTime.Now;

                            txtItem_Number = "ItemNum";
                            txtReference_Number = "RefNum";
                            txtStocking_UOM = "UOMS";
                            txtItem_Description = "ItemDescription";
                            txtQuantity = "Qt";
                            txtCustomer_Number = "CustNum";
                            txtCustomer_Name = "CustName";
                            txtPO_Number = "PurchaseNum";
                            txtDate = "ToDate";
                            txtTime = "ToTime";
                            custItemNo = "CustItemNo";

                            string wordTemplate = Path.Combine(ConfigurationManager.AppSettings["LabelTempTemplateFolder"], "721-LX Template " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".docx");
                            File.Copy(originalTemplateWordDocument, wordTemplate);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "New Work Template Created: " + wordTemplate);
                            using (WordDocument documents = new WordDocument())
                            {
                                documents.Open(wordTemplate);
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Opened New Template");

                                documents.Replace(txtItem_Number, Item_Number, false, true);
                                documents.Replace(txtReference_Number, Reference_Number, false, true);
                                documents.Replace(txtStocking_UOM, Stocking_UOM, false, true);
                                documents.Replace(txtItem_Description, Item_Description, false, true);
                                documents.Replace(txtQuantity, Quantity, false, true);
                                documents.Replace(txtCustomer_Number, Customer_Number, false, true);
                                documents.Replace(txtCustomer_Name, Customer_Name, false, true);
                                documents.Replace(custItemNo, customerItemNumber, false, true);
                                documents.Replace(txtPO_Number, PO_Number, false, true);
                                documents.Replace(txtDate, datetime.ToString("dd/MM/yyyy"), false, true);
                                documents.Replace(txtTime, datetime.ToString("HH:mm:ss"), false, true);
                                documents.Save(wordTemplate);
                                documents.Close();
                            }

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Saved and Closed Template");

                            string newPDFFileName = ConfigurationManager.AppSettings["LabelTempTemplateFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "721-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf";
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "New PDF Document Created: " + newPDFFileName);
                            using (DocToPDFConverter converter = new DocToPDFConverter())
                            {
                                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Convert: " + wordTemplate + " To: " + newPDFFileName);
                                using (PdfDocument pdfDocument = converter.ConvertToPDF(wordTemplate))
                                {
                                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Converted and Saved PDF Document");

                                    try
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Start Barcode Insert");
                                        PdfPage page = pdfDocument.Pages[0];
                                        PdfCode39Barcode barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 12;
                                        barcode.Size = new SizeF(115, 12);
                                        barcode.Text = Item_Number;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(90, 75));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 12;
                                        barcode.Size = new SizeF(145, 12);
                                        barcode.Text = customerItemNumber;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(90, 133));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 12;
                                        barcode.Size = new SizeF(90, 12);
                                        barcode.Text = PO_Number;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(290, 53));

                                        barcode = new PdfCode39Barcode();
                                        barcode.BarHeight = 12;
                                        barcode.Size = new SizeF(80, 12);
                                        barcode.Text = Quantity;
                                        barcode.TextDisplayLocation = TextLocation.None;
                                        barcode.Draw(page, new PointF(298, 220));

                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Inserted Barcodes");
                                    }
                                    catch (Exception ex)
                                    {
                                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Failed to Insert Barcode - Error " + ex.ToString());
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

                            string outputPDFFile = ConfigurationManager.AppSettings["LabelOutputFolder"] + "(" + Printer_IP.Replace(@"\", "") + ")" + "721-LX Converted PDF " + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "(" + totalNumberOdPages.ToString() + ")" + ".pdf";
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
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Deleted: " + wordTemplate);
                            File.Delete(newPDFFileName);
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Deleted: " + newPDFFileName);
                        }
                        else
                        {
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "No Template Found");
                        }
                    }
                    catch (Exception ex)
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Failed to Process - Error " + ex.ToString());
                    }

                    try
                    {
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX", "Finished Processing Label");
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "721-LX");
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
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "721-LX FAILED", lofFileData[0].Name + " : " + ex.ToString());
                        csFileInputEngine.MoveFileToArchive(lofFileData[0], "721-LX FAILED");
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
                        RearrangeColumnEnd = 41,
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
                        RearrangeColumnEnd = 42,
                        RearrangeRowNumber = 6
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
                        RearrangeColumnEnd = 32,
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
                        RearrangeColumnEnd = 59,
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
                        RearrangeColumnEnd = 15,
                        RearrangeRowNumber = 15
                    };
                }
            }


            public RearrangeSetting Customer_Number
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 58,
                        RearrangeRowNumber = 18
                    };
                }
            }
            public RearrangeSetting Customer_Name
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 71,
                        RearrangeRowNumber = 19
                    };
                }
            }
            public RearrangeSetting CustomerItemNumber
            {
                get
                {
                    return new RearrangeSetting()
                    {
                        RearrangeColumnStart = 25,
                        RearrangeColumnEnd = 57,
                        RearrangeRowNumber = 16
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
                        RearrangeRowNumber = 21
                    };
                }
            }
        }

    }
}
