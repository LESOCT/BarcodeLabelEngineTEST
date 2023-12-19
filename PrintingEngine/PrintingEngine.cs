using Microsoft.Win32.SafeHandles;
using Syncfusion.Windows.Forms.PdfViewer;
using Syncfusion.Windows.PdfViewer;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BarcodeLabelSoftware
{
    public class PrintingEngine : IDisposable
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
                LogEngine logEngine = new LogEngine(); logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Tasks", "Disposed Print Engine"); handle.Dispose();
            }

            disposed = true;
        }

        public List<FileInfo> printerQueue { get; set; }
        public Dictionary<string, Task> lofPrintTasks = new Dictionary<string, Task>();
        public async Task LabelRouting()
        {
            await Task.Run(() => Start());
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void Start()
        {
            //PrinterControl printerQueue1 = null;
            //PrinterControl printerQueue2 = null;
            //PrinterControl printerQueue3 = null;
            //PrinterControl printerQueue4 = null;
            //PrinterControl printerQueue5 = null;
            //PrinterControl printerQueue6 = null;
            //PrinterControl printerQueue7 = null;
            //PrinterControl printerQueue8 = null;
            //PrinterControl printerQueue9 = null;
            //PrinterControl printerQueue10 = null;

            bool endless = true;
            while (endless)
            {
                try
                {
                    if(printerQueue.Count > 0)
                    {
                        FileInfo tempLabel = printerQueue[0];
                        if (tempLabel.Exists)
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
                        printerQueue.RemoveAt(0);
                    }
                }
                catch (Exception ex)
                {
                    if (printerQueue.Count > 0)
                    {
                        printerQueue.RemoveAt(0);
                    }
                    LogEngine logEngine = new LogEngine();
                    logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Failed Printer Log", "Failed to Print File - Error " + ex.ToString());
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }




            //try
            //{
            //    LogEngine logEngine = new LogEngine();
            //    if (printerQueue.Count > 0)
            //    {
            //        try
            //        {
            //            if (printerQueue1 == null)
            //            {
            //                printerQueue1 = new PrinterControl();
            //                printerQueue1.tempLabel = printerQueue[0];
            //                Task printTask = printerQueue1.PrintDocument();
            //                lofPrintTasks.Add("1", printTask);
            //                printerQueue.RemoveAt(0);
            //            }
            //            //else if (printerQueue2 == null)
            //            //{
            //            //    printerQueue2 = new PrinterControl();
            //            //    printerQueue2.tempLabel = printerQueue[0];
            //            //    Task printTask = printerQueue2.PrintDocument();
            //            //    lofPrintTasks.Add("2", printTask);
            //            //    printerQueue.RemoveAt(0);
            //            //}
            //            //else if (printerQueue3 == null)
            //            //{
            //            //    printerQueue3 = new PrinterControl();
            //            //    printerQueue3.tempLabel = printerQueue[0];
            //            //    Task printTask = printerQueue3.PrintDocument();
            //            //    lofPrintTasks.Add("3", printTask);
            //            //    printerQueue.RemoveAt(0);
            //            //}
            //            //else if (printerQueue4 == null)
            //            //{
            //            //    printerQueue4 = new PrinterControl();
            //            //    printerQueue4.tempLabel = printerQueue[0];
            //            //    Task printTask = printerQueue4.PrintDocument();
            //            //    lofPrintTasks.Add("4", printTask);
            //            //    printerQueue.RemoveAt(0);
            //            //}
            //            //else if (printerQueue5 == null)
            //            //{
            //            //    printerQueue5 = new PrinterControl();
            //            //    printerQueue5.tempLabel = printerQueue[0];
            //            //    Task printTask = printerQueue5.PrintDocument();
            //            //    lofPrintTasks.Add("5", printTask);
            //            //    printerQueue.RemoveAt(0);
            //            //}
            //            //else if (printerQueue6 == null)
            //            //{
            //            //    printerQueue6 = new PrinterControl();
            //            //    printerQueue6.tempLabel = printerQueue[0];
            //            //    Task printTask = printerQueue6.PrintDocument();
            //            //    lofPrintTasks.Add("6", printTask);
            //            //    printerQueue.RemoveAt(0);
            //            //}
            //            //else if (printerQueue7 == null)
            //            //{
            //            //    printerQueue7 = new PrinterControl();
            //            //    printerQueue7.tempLabel = printerQueue[0];
            //            //    Task printTask = printerQueue7.PrintDocument();
            //            //    lofPrintTasks.Add("7", printTask);
            //            //    printerQueue.RemoveAt(0);
            //            //}
            //            //else if (printerQueue8 == null)
            //            //{
            //            //    printerQueue8 = new PrinterControl();
            //            //    printerQueue8.tempLabel = printerQueue[0];
            //            //    Task printTask = printerQueue8.PrintDocument();
            //            //    lofPrintTasks.Add("8", printTask);
            //            //    printerQueue.RemoveAt(0);
            //            //}
            //            //else if (printerQueue9 == null)
            //            //{
            //            //    printerQueue9 = new PrinterControl();
            //            //    printerQueue9.tempLabel = printerQueue[0];
            //            //    Task printTask = printerQueue9.PrintDocument();
            //            //    lofPrintTasks.Add("9", printTask);
            //            //    printerQueue.RemoveAt(0);
            //            //}
            //            //else if (printerQueue10 == null)
            //            //{
            //            //    printerQueue10 = new PrinterControl();
            //            //    printerQueue10.tempLabel = printerQueue[0];
            //            //    Task printTask = printerQueue10.PrintDocument();
            //            //    lofPrintTasks.Add("10", printTask);
            //            //    printerQueue.RemoveAt(0);
            //            //}
            //        }
            //        catch (Exception ex)
            //        {
            //            try
            //            {
            //                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Failed Task", ex.ToString());
            //                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Failed Printer Log", "Failed to Print: " + printerQueue[0].FullName + " - Error " + ex.ToString());
            //                File.SetAttributes(printerQueue[0].FullName, FileAttributes.Normal);
            //                File.Delete(printerQueue[0].FullName);
            //            }
            //            catch
            //            {
            //                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Failed Printer Log", "Failed to Remove File - Error " + ex.ToString());
            //            }
            //        }
            //    }
            //    try
            //    {
            //        List<string> lofCompletedTasks = new List<string>();
            //        foreach (KeyValuePair<string, Task> task in lofPrintTasks)
            //        {
            //            if (task.Value.IsCompleted == true || task.Value.Status == TaskStatus.RanToCompletion || task.Value.Status == TaskStatus.Faulted)
            //            {
            //                task.Value.Dispose();
            //                lofCompletedTasks.Add(task.Key);
            //                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Tasks", "Disposed Printer Queue Task: " + task.Key);

            //                GC.Collect();
            //                GC.WaitForPendingFinalizers();
            //                GC.Collect();
            //            }
            //        }

            //        foreach (string task in lofCompletedTasks)
            //        {
            //            if (task == "1")
            //            {
            //                printerQueue1.Dispose();
            //                printerQueue1 = null;
            //            }
            //            //else if (task == "2")
            //            //{
            //            //    printerQueue2.Dispose();
            //            //    printerQueue2 = null;
            //            //}
            //            //else if (task == "3")
            //            //{
            //            //    printerQueue3.Dispose();
            //            //    printerQueue3 = null;
            //            //}
            //            //else if (task == "4")
            //            //{
            //            //    printerQueue4.Dispose();
            //            //    printerQueue4 = null;
            //            //}
            //            //else if (task == "5")
            //            //{
            //            //    printerQueue5.Dispose();
            //            //    printerQueue5 = null;
            //            //}
            //            //else if (task == "6")
            //            //{
            //            //    printerQueue6.Dispose();
            //            //    printerQueue6 = null;
            //            //}
            //            //else if (task == "7")
            //            //{
            //            //    printerQueue7.Dispose();
            //            //    printerQueue7 = null;
            //            //}
            //            //else if (task == "8")
            //            //{
            //            //    printerQueue8.Dispose();
            //            //    printerQueue8 = null;
            //            //}
            //            //else if (task == "9")
            //            //{
            //            //    printerQueue9.Dispose();
            //            //    printerQueue9 = null;
            //            //}
            //            //else if (task == "10")
            //            //{
            //            //    printerQueue10.Dispose();
            //            //    printerQueue10 = null;
            //            //}
            //            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Tasks", "Disposed Printer Queue Class: " + task);
            //            lofPrintTasks.Remove(task);

            //            GC.Collect();
            //            GC.WaitForPendingFinalizers();
            //            GC.Collect();
            //        }
            //        lofCompletedTasks.Clear();
            //    }
            //    catch (Exception ex)
            //    {
            //        try
            //        {
            //            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Failed Task", ex.ToString());
            //        }
            //        catch
            //        {

            //        }
            //    }

            //    if (printerQueue.Count == 0 && lofPrintTasks.Count == 0)
            //    {
            //        break;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    try
            //    {
            //        LogEngine logEngine = new LogEngine();
            //        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "PRINT FAILED", printerQueue[0].Name + " : " + ex.ToString());
            //        printerQueue.RemoveAt(0);
            //    }
            //    catch
            //    {

            //    }
            //}
        }
    }
}


