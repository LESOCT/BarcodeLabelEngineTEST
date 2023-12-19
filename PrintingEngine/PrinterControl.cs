using Microsoft.Win32.SafeHandles;
using Syncfusion.Windows.Forms.PdfViewer;
using Syncfusion.Windows.PdfViewer;
using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BarcodeLabelSoftware
{
    public class PrinterControl : IDisposable
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
                LogEngine logEngine = new LogEngine(); logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Tasks", "Disposed Printer Control Engine"); handle.Dispose();
            }

            disposed = true;
        }

        public FileInfo tempLabel { get; set; }

        public async Task PrintDocument()
        {
            await Task.Run(() => Start());
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void Start()
        {
            try
            {
                DirectoryInfo printerTempFolder = new DirectoryInfo(ConfigurationManager.AppSettings["LabelPrinterTempFolder"]);
                FileInfo label = new FileInfo(Path.Combine(printerTempFolder.FullName, tempLabel.Name));
                File.Move(tempLabel.FullName, label.FullName);
                string printerIP = label.Name.Substring(1, label.Name.IndexOf(")") - 1);
                string tempNumberOfCopies = label.Name.Substring(label.Name.LastIndexOf("(") + 1);
                int numberOfCopies = Convert.ToInt32(tempNumberOfCopies.Substring(0, tempNumberOfCopies.LastIndexOf(")")));
                if (tempLabel.Name.Contains("702-LX") || tempLabel.Name.Contains("703-LX") || tempLabel.Name.Contains("704-LX") || tempLabel.Name.Contains("706-LX") || tempLabel.Name.Contains("707-LX") || tempLabel.Name.Contains("708-LX"))
                {
                    if (tempLabel.Name.Contains("706-LX") || tempLabel.Name.Contains("708-LX"))
                    {
                        printerIP = printerIP + "_1";
                    }

                    try
                    {
                        for (int i = 1; i <= numberOfCopies; i++)
                        {
                            ProcessStartInfo psInfo = new ProcessStartInfo();
                            psInfo.FileName = ConfigurationManager.AppSettings["FoxitReaderLocation"];
                            psInfo.Arguments = String.Format("/t \"{0}\" \"{1}\"",
                                label.FullName,
                                printerIP);
                            psInfo.WindowStyle = ProcessWindowStyle.Hidden;
                            psInfo.CreateNoWindow = true;
                            psInfo.UseShellExecute = true;
                            Process process = Process.Start(psInfo);
                            process.WaitForExit(10000);
                            if (!process.HasExited)
                            {
                                process.Kill();
                                process.Dispose();
                            } 
                        }  
                    }
                    catch
                    {

                    }
                }
                else
                {
                    using (PdfViewerControl viewer = new PdfViewerControl())
                    {
                        viewer.PrinterSettings.PageSize = PdfViewerPrintSize.Fit;
                        viewer.Load(label.FullName);

                        using (PrintDialog dialog = new PrintDialog())
                        {
                            dialog.AllowPrintToFile = true;
                            dialog.AllowSomePages = true;
                            dialog.AllowCurrentPage = true;
                            dialog.Document = viewer.PrintDocument;
                            dialog.Document.PrinterSettings.PrinterName = printerIP;
                            dialog.Document.DocumentName = label.Name;
                            dialog.Document.PrinterSettings.Copies = (short)numberOfCopies;
                            dialog.Document.Print();
                            
                            dialog.Document.Dispose();
                            dialog.Dispose();
                        }

                        viewer.Unload();
                        viewer.Dispose();
                    }  
                }

                LogEngine logEngine = new LogEngine();
                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Printer Log", "Successfully Printed: " + tempLabel.FullName + " To: " + printerIP);
            }
            catch(Exception ex)
            {
                LogEngine logEngine = new LogEngine();
                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Failed Printer Log", "Failed to Print File - Error " + ex.ToString());
            }
        }
    }
}
