using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace BarcodeLabelSoftware
{
    public class FileEngine
    {
        public List<FileInfo> GetAllFiles(DirectoryInfo inputDirectory)
        {
            try
            {
                return inputDirectory.GetFiles("*.*").ToList();
            }
            catch
            {
                return new List<FileInfo>();
            }
        }

        public List<FileInfo> GetAllPDFFiles(DirectoryInfo inputDirectory)
        {
            try
            {
                return inputDirectory.GetFiles("*.pdf").ToList();
            }
            catch
            {
                return new List<FileInfo>();
            }
        }

        public void CleanPrinterQueue()
        {
            LogEngine logEngine = new LogEngine();
            DirectoryInfo printerTempFolder = new DirectoryInfo(ConfigurationManager.AppSettings["LabelPrinterTempFolder"]);
            try
            {
                if (!printerTempFolder.Exists)
                {
                    printerTempFolder.Create();
                }
            }
            catch
            {
                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "File Log", "Unable to create: " + ConfigurationManager.AppSettings["LabelPrinterTempFolder"]);
            }

            string fullFileName = "";
            try
            {
                List<FileInfo> files = printerTempFolder.GetFiles("*.pdf").ToList();
                foreach (FileInfo file in files)
                {
                    if (file.LastAccessTime < DateTime.Now.AddMinutes(-10))
                    {
                        fullFileName = file.FullName;
                        File.SetAttributes(fullFileName, FileAttributes.Normal);
                        File.Delete(fullFileName);
                        logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "File Log", "Successfully Removed File: " + fullFileName);
                    } 
                }
            }
            catch(Exception ex)
            {
                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "File Log", "Failed To Removed File: " + fullFileName + " Error: " + ex.ToString());
            }
        }

        public List<FileInfo> TransferFileForProcesssing(List<FileInfo> lofInputFiles)
        {
            List<FileInfo> lofTempProcessingFiles = new List<FileInfo>();
            DirectoryInfo tempProcessingDirectory = new DirectoryInfo(ConfigurationManager.AppSettings["LabelTempFolder"]);
            if(!tempProcessingDirectory.Exists)
            {
                try { tempProcessingDirectory.Create(); } catch { }
            }
            foreach(FileInfo inputFile in lofInputFiles)
            {
                try
                {
                    string newFileName = Path.Combine(tempProcessingDirectory.FullName, inputFile.Name);
                    if (File.Exists(newFileName))
                    {
                        bool uniqueNameFound = false;
                        int count = 1;
                        while (!uniqueNameFound)
                        {
                            newFileName = Path.Combine(tempProcessingDirectory.FullName, Path.GetFileNameWithoutExtension(inputFile.FullName) + "(" + count + ")" + Path.GetExtension(inputFile.FullName));
                            if (File.Exists(newFileName))
                            {
                                count++;
                            }
                            else
                            {
                                uniqueNameFound = true;
                            }
                        }
                    }

                    //File.SetAttributes(inputFile.FullName, FileAttributes.Normal);
                    File.Move(inputFile.FullName, newFileName);

                    lofTempProcessingFiles.Add(new FileInfo(newFileName));
                }
                catch(Exception ex)
                {

                }
            }
            return lofTempProcessingFiles;
        }

        public void MoveFileToArchive(FileInfo processedFile, string archiveName)
        {
            try
            {
                DirectoryInfo archiveDirectory = new DirectoryInfo(Path.Combine(ConfigurationManager.AppSettings["LabelArchiveFolder"] + @"\" + DateTime.Now.ToString("yyyy-MM-dd"), archiveName));
                if (!archiveDirectory.Exists)
                {
                    archiveDirectory.Create();
                }
                string newFileName = Path.Combine(archiveDirectory.FullName, processedFile.FullName);
                if (File.Exists(newFileName))
                {
                    bool uniqueNameFound = false;
                    int count = 1;
                    while (!uniqueNameFound)
                    {
                        newFileName = Path.Combine(archiveDirectory.FullName, Path.GetFileNameWithoutExtension(processedFile.FullName) + "(" + count + ")" + Path.GetExtension(processedFile.FullName));
                        if (File.Exists(newFileName))
                        {
                            count++;
                        }
                        else
                        {
                            uniqueNameFound = true;
                        }
                    }
                }
                File.Move(processedFile.FullName, newFileName);
            }
            catch
            {

            }
        }
    }
}
