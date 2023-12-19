using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace BarcodeLabelSoftware
{
    class Program
    {
        static void Main()
        {
            if (System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1)
            {
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("MTk0Njg5QDMxMzcyZTMzMmUzMFdlREpRQlJKd1QzcGV0VFdVOG9taEZIOEIyenU5L095M3J1SXlQWEI1clE9;MTk0NjkwQDMxMzcyZTMzMmUzMGNQWHVhdUwzbDE1RlV5d0dkUkkwUUUzc3lRa1ZQQmVBUHRLTW1tWUQ3ajQ9;MTk0NjkxQDMxMzcyZTMzMmUzMFNiSXF3V0hxUVZCK3UvUndtamZWY3FyRVFPMDFxZ0c0aVJxdGhxVXVIQVU9;MTk0NjkyQDMxMzcyZTMzMmUzMEpmeXZvZzI3YmRKWUI5Y2t4bVlQM1ZBUEM0TDg4eE5obWJsdDlFR0hHdmc9;MTk0NjkzQDMxMzcyZTMzMmUzMFhIR3k2ZWJFSzNOMHZ4djZPeFpZMHZQeGIxdXU1NTc4bjFCWWZmZXdlckk9;MTk0Njk0QDMxMzcyZTMzMmUzMFBHUStoeDdvQ1NqYmtpQ1R1NUEraDYyTG9jQWNxaEsrQlhNd1Fvd3NScm89;MTk0Njk1QDMxMzcyZTMzMmUzMFRhVUUrZnFUOFFVMWM0MmRaTmpKejlKSkZ3Y1ZTVmVxamJSdnhuVEY0bUE9;MTk0Njk2QDMxMzcyZTMzMmUzMGdJMDFCTURiQkMrZGlOWnQyY1hsUVY2K0x6T2taNDk0Um8ybm1zdGtYa2s9;MTk0Njk3QDMxMzcyZTMzMmUzMG1LeCtveUhybVhPOFFTdUxDeGp2OHhNYmhKdVk0VWF2S24yVVo2M3FnRUU9;NT8mJyc2IWhiZH1gfWN9YGdoYmF8YGJ8ampqanNiYmlmamlmanMDHmg5PD0yJzsyPX0gPCYnOyQ8PDcTPzYgfTA8fSky;MTk0Njk4QDMxMzcyZTMzMmUzMGFMZ1lLelFrbTlHZ1NkNlpIWTMxcVZXOGlJaFpIWUhSYnY4N1FoUFBKR1k9");


            DirectoryInfo inputDirectory = new DirectoryInfo(ConfigurationManager.AppSettings["LabelInputFolder"]);
            bool endless = true;
            LabelRoutingEngine csLabelRoutingEngine = null;
            //PrintingEngine csPrintingEngine = null;
            //Task printingTask = null;
            Task routingTask = null;
            while (endless)
            {
                FileEngine csFileInputEngine = new FileEngine();
                //List<FileInfo> printerQueue = csFileInputEngine.GetAllPDFFiles(new DirectoryInfo(ConfigurationManager.AppSettings["LabelOutputFolder"]));
                //csFileInputEngine.CleanPrinterQueue();

                //if (csPrintingEngine == null)
                //{
                //    csPrintingEngine = new PrintingEngine();
                //    csPrintingEngine.printerQueue = new List<FileInfo>();
                //}
                
                //foreach(FileInfo job in printerQueue)
                //{
                //    if (!csPrintingEngine.printerQueue.Any(a => a.FullName == job.FullName) && job.Exists)
                //    {
                //        csPrintingEngine.printerQueue.Add(job);
                //    }
                //}
                
                //printerQueue.Clear();
                //if (csPrintingEngine.printerQueue.Count > 0)
                //{
                //    if (printingTask == null)
                //    {
                //        printingTask = csPrintingEngine.LabelRouting();
                //    }
                //    else if (printingTask.IsCompleted == true &&
                //       printingTask.Status != TaskStatus.Running &&
                //       printingTask.Status != TaskStatus.WaitingToRun &&
                //       printingTask.Status != TaskStatus.WaitingForActivation)
                //    {
                //        printingTask.Dispose();
                //        printingTask = csPrintingEngine.LabelRouting();
                //    }
                //}
                //else
                //{
                //    if (printingTask != null && (printingTask.IsCompleted == true || printingTask.Status == TaskStatus.RanToCompletion || printingTask.Status == TaskStatus.Faulted) && csPrintingEngine.lofPrintTasks.Count == 0)
                //    {
                //        printingTask.Dispose();
                //        printingTask = null;
                //        csPrintingEngine.Dispose();
                //        csPrintingEngine = null;
                //        GC.Collect();
                //        GC.WaitForPendingFinalizers();
                //        GC.Collect();
                //    }
                //}

                if (inputDirectory.Exists)
                {
                    List<FileInfo> lofInputFiles = csFileInputEngine.GetAllFiles(inputDirectory);
                    if(lofInputFiles.Count > 0)
                    {
                        if (csLabelRoutingEngine == null)
                        {
                            csLabelRoutingEngine = new LabelRoutingEngine();
                            csLabelRoutingEngine.routingFiles = new List<FileInfo>();
                        }

                        csLabelRoutingEngine.routingFiles.AddRange(csFileInputEngine.TransferFileForProcesssing(lofInputFiles));
                        lofInputFiles.Clear();
                        if (csLabelRoutingEngine.routingFiles.Count > 0)
                        {
                            if (routingTask == null)
                            {
                                routingTask = csLabelRoutingEngine.LabelRouting();
                            }
                            else if (routingTask.IsCompleted == true &&
                               routingTask.Status != TaskStatus.Running &&
                               routingTask.Status != TaskStatus.WaitingToRun &&
                               routingTask.Status != TaskStatus.WaitingForActivation)
                            {
                                routingTask.Dispose();
                                routingTask = csLabelRoutingEngine.LabelRouting();
                            }
                        }
                        else
                        {
                            if (routingTask != null && (routingTask.IsCompleted == true || routingTask.Status == TaskStatus.RanToCompletion || routingTask.Status == TaskStatus.Faulted) && csLabelRoutingEngine.lofTasks.Count == 0)
                            {
                                routingTask.Dispose();
                                routingTask = null;
                                csLabelRoutingEngine.Dispose();
                                csLabelRoutingEngine = null;
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                GC.Collect();
                            }
                        }
                    }
                    else
                    {
                        if (routingTask != null && (routingTask.IsCompleted == true || routingTask.Status == TaskStatus.RanToCompletion || routingTask.Status == TaskStatus.Faulted) && csLabelRoutingEngine.lofTasks.Count == 0)
                        {
                            routingTask.Dispose();
                            routingTask = null;
                            csLabelRoutingEngine.Dispose();
                            csLabelRoutingEngine = null;
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                        }
                    }
                    Thread.Sleep(2000);
                }
                else
                {
                    try
                    {
                        inputDirectory.Create();
                    }
                    catch
                    {
                        break;
                    }
                }
            }
        }
    }
}
