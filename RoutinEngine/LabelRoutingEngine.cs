using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BarcodeLabelSoftware
{
    public class LabelRoutingEngine : IDisposable
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
                LogEngine logEngine = new LogEngine();
                logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Tasks", "Disposed Routing Task");
                handle.Dispose();
            }

            disposed = true;
        }
    
        public List<FileInfo> routingFiles { get; set; }
        public async Task LabelRouting()
        {
            await Task.Run(() => Start());
        }

        public Dictionary<string, Task> lofTasks = new Dictionary<string, Task>();
        private void Start()
         {
            LX704LabelEngine csLX704LabelEngine = null;
            LX702LabelEngine csLX702LabelEngine = null;
            LX703LabelEngine csLX703LabelEngine = null;
            LX706LabelEngine csLX706LabelEngine = null;
            LX707LabelEngine csLX707LabelEngine = null;
            LX708LabelEngine csLX708LabelEngine = null;
            LX710LabelEngine csLX710LabelEngine = null;
            LX711LabelEngine csLX711LabelEngine = null;
            LX712LabelEngine csLX712LabelEngine = null;
            LX713LabelEngine csLX713LabelEngine = null;
            LX717LabelEngine csLX717LabelEngine = null;
            LX718LabelEngine csLX718LabelEngine = null;
            LX719LabelEngine csLX719LabelEngine = null;
            LX720LabelEngine csLX720LabelEngine = null;
            LX721LabelEngine csLX721LabelEngine = null;
            LX723LabelEngine csLX723LabelEngine = null;
            LX724LabelEngine csLX724LabelEngine = null;

            bool endless = true;
            while (endless)
            {
                while (routingFiles.Count > 0)
                {
                    try
                    {
                        string[] lofLines = File.ReadAllText(routingFiles[0].FullName).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                        LabelCriteriaSettings documentCriteria = new LabelCriteriaSettings();
                        bool foundMatch = false;
                        foreach (List<CriteriaSetting> lofCriteriaSettings in documentCriteria.lofLabelCriteriaSettings)
                        {
                            int numberOfMatches = 0;
                            foreach (CriteriaSetting criteriaSetting in lofCriteriaSettings)
                            {
                                string textContent = "";

                                if (criteriaSetting.RoutingLabel == "720-LX")//DONT DELETE // COMPLEX CRITERIA
                                {
                                    string labelType = "";
                                    try
                                    {
                                        labelType = lofLines[1].Substring(25, 15).Trim();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            labelType = lofLines[1].Substring(25).Trim();
                                        }
                                        catch
                                        {

                                        }
                                    }

                                    string NegativeCriteria1 = "";
                                    try
                                    {
                                        NegativeCriteria1 = lofLines[19].Substring(24, 4).Trim();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            NegativeCriteria1 = lofLines[19].Substring(24).Trim();
                                        }
                                        catch
                                        {

                                        }
                                    }

                                    string NegativeCriteria2 = "";
                                    try
                                    {
                                        NegativeCriteria2 = lofLines[7].Substring(25, 3).Trim();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            NegativeCriteria2 = lofLines[7].Substring(25).Trim();
                                        }
                                        catch
                                        {

                                        }
                                    }

                                    string NegativeCriteria3 = "";
                                    try
                                    {
                                        NegativeCriteria3 = lofLines[7].Substring(25, 3).Trim();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            NegativeCriteria3 = lofLines[7].Substring(25).Trim();
                                        }
                                        catch
                                        {

                                        }
                                    }

                                    string NegativeCriteria4 = "";
                                    try
                                    {
                                        NegativeCriteria4 = lofLines[19].Substring(24, 4).Trim();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            NegativeCriteria4 = lofLines[19].Substring(24).Trim();
                                        }
                                        catch
                                        {

                                        }
                                    }

                                    string NegativeCriteria5 = "";
                                    try
                                    {
                                        NegativeCriteria5 = lofLines[13].Substring(25, 6).Trim();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            NegativeCriteria5 = lofLines[13].Substring(25).Trim();
                                        }
                                        catch
                                        {

                                        }
                                    }

                                    string NegativeCriteria6 = "";
                                    try
                                    {
                                        NegativeCriteria6 = lofLines[2].Substring(27, 13).Trim();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            NegativeCriteria6 = lofLines[2].Substring(27).Trim();
                                        }
                                        catch
                                        {

                                        }
                                    }

                                    string NegativeCriteria7 = "";
                                    try
                                    {
                                        NegativeCriteria7 = lofLines[6].Substring(25, 2).Trim();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            NegativeCriteria7 = lofLines[6].Substring(25).Trim();
                                        }
                                        catch
                                        {

                                        }
                                    }

                                    if (labelType == "Transfer Booked" && (NegativeCriteria1 != "=NW" && NegativeCriteria2 != "J1" && NegativeCriteria3 != "S1" && NegativeCriteria4 != "=TG" && NegativeCriteria5 != "mldp25" && NegativeCriteria6 != "192.168.4.223" && NegativeCriteria7 != "L4"))
                                    {
                                        numberOfMatches++;
                                    }
                                }
                                else
                                {
                                    
                                    try
                                    {
                                        textContent = lofLines[criteriaSetting.CriteriaRowNumber - 1].Substring(criteriaSetting.CriteriaColumnStart, criteriaSetting.CriteriaColumnEnd).Trim();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            textContent = lofLines[criteriaSetting.CriteriaRowNumber - 1].Substring(criteriaSetting.CriteriaColumnStart).Trim();
                                        }
                                        catch
                                        {

                                        }
                                    }

                                    if (textContent == criteriaSetting.MatchString)
                                    {
                                        numberOfMatches++;
                                    }
                                }
                            }
                            

                            if (numberOfMatches == lofCriteriaSettings.Count)
                            {
                                foundMatch = true;
                                switch (lofCriteriaSettings[0].RoutingLabel)
                                {
                                    case "702-LX":
                                        if (lofTasks.ContainsKey("702-LX"))
                                        {
                                            if (csLX702LabelEngine == null)
                                            {
                                                csLX702LabelEngine = new LX702LabelEngine();
                                                csLX702LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX702LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["702-LX"].IsCompleted == true &&
                                                lofTasks["702-LX"].Status != TaskStatus.Running &&
                                                lofTasks["702-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["702-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["702-LX"].Dispose();
                                                lofTasks["702-LX"] = csLX702LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX702LabelEngine = new LX702LabelEngine();
                                            csLX702LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX702LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx702LabelEngine = csLX702LabelEngine.GenerateLabel();
                                            lofTasks.Add("702-LX", lx702LabelEngine);
                                        }
                                        break;
                                    case "703-LX":
                                        if (lofTasks.ContainsKey("703-LX"))
                                        {
                                            if (csLX703LabelEngine == null)
                                            {
                                                csLX703LabelEngine = new LX703LabelEngine();
                                                csLX703LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX703LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["703-LX"].IsCompleted == true &&
                                                lofTasks["703-LX"].Status != TaskStatus.Running &&
                                                lofTasks["703-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["703-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["703-LX"].Dispose();
                                                lofTasks["703-LX"] = csLX703LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX703LabelEngine = new LX703LabelEngine();
                                            csLX703LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX703LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx703LabelEngine = csLX703LabelEngine.GenerateLabel();
                                            lofTasks.Add("703-LX", lx703LabelEngine);
                                        }
                                        break;
                                    case "704-LX":
                                        if (lofTasks.ContainsKey("704-LX"))
                                        {
                                            if (csLX704LabelEngine == null)
                                            {
                                                csLX704LabelEngine = new LX704LabelEngine();
                                                csLX704LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX704LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["704-LX"].IsCompleted == true &&
                                                lofTasks["704-LX"].Status != TaskStatus.Running &&
                                                lofTasks["704-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["704-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["704-LX"].Dispose();
                                                lofTasks["704-LX"] = csLX704LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX704LabelEngine = new LX704LabelEngine();
                                            csLX704LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX704LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx704LabelEngine = csLX704LabelEngine.GenerateLabel();
                                            lofTasks.Add("704-LX", lx704LabelEngine);
                                        }
                                        break;
                                    case "706-LX":
                                        if (lofTasks.ContainsKey("706-LX"))
                                        {
                                            if (csLX706LabelEngine == null)
                                            {
                                                csLX706LabelEngine = new LX706LabelEngine();
                                                csLX706LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX706LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["706-LX"].IsCompleted == true &&
                                                lofTasks["706-LX"].Status != TaskStatus.Running &&
                                                lofTasks["706-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["706-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["706-LX"].Dispose();
                                                lofTasks["706-LX"] = csLX706LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX706LabelEngine = new LX706LabelEngine();
                                            csLX706LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX706LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx706LabelEngine = csLX706LabelEngine.GenerateLabel();
                                            lofTasks.Add("706-LX", lx706LabelEngine);
                                        }
                                        break;
                                    case "707-LX":
                                        if (lofTasks.ContainsKey("707-LX"))
                                        {
                                            if (csLX707LabelEngine == null)
                                            {
                                                csLX707LabelEngine = new LX707LabelEngine();
                                                csLX707LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX707LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["707-LX"].IsCompleted == true &&
                                                lofTasks["707-LX"].Status != TaskStatus.Running &&
                                                lofTasks["707-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["707-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["707-LX"].Dispose();
                                                lofTasks["707-LX"] = csLX707LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX707LabelEngine = new LX707LabelEngine();
                                            csLX707LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX707LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx707LabelEngine = csLX707LabelEngine.GenerateLabel();
                                            lofTasks.Add("707-LX", lx707LabelEngine);
                                        }
                                        break;
                                    case "708-LX":
                                        if (lofTasks.ContainsKey("708-LX"))
                                        {
                                            if (csLX708LabelEngine == null)
                                            {
                                                csLX708LabelEngine = new LX708LabelEngine();
                                                csLX708LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX708LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["708-LX"].IsCompleted == true &&
                                                lofTasks["708-LX"].Status != TaskStatus.Running &&
                                                lofTasks["708-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["708-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["708-LX"].Dispose();
                                                lofTasks["708-LX"] = csLX708LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX708LabelEngine = new LX708LabelEngine();
                                            csLX708LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX708LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx708LabelEngine = csLX708LabelEngine.GenerateLabel();
                                            lofTasks.Add("708-LX", lx708LabelEngine);
                                        }
                                        break;
                                    case "710-LX":
                                        if (lofTasks.ContainsKey("710-LX"))
                                        {
                                            if (csLX710LabelEngine == null)
                                            {
                                                csLX710LabelEngine = new LX710LabelEngine();
                                                csLX710LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX710LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["710-LX"].IsCompleted == true &&
                                                lofTasks["710-LX"].Status != TaskStatus.Running &&
                                                lofTasks["710-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["710-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["710-LX"].Dispose();
                                                lofTasks["710-LX"] = csLX710LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX710LabelEngine = new LX710LabelEngine();
                                            csLX710LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX710LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx710LabelEngine = csLX710LabelEngine.GenerateLabel();
                                            lofTasks.Add("710-LX", lx710LabelEngine);
                                        }
                                        break;
                                    case "711-LX":
                                        if (lofTasks.ContainsKey("711-LX"))
                                        {
                                            if (csLX711LabelEngine == null)
                                            {
                                                csLX711LabelEngine = new LX711LabelEngine();
                                                csLX711LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX711LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["711-LX"].IsCompleted == true &&
                                                lofTasks["711-LX"].Status != TaskStatus.Running &&
                                                lofTasks["711-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["711-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["711-LX"].Dispose();
                                                lofTasks["711-LX"] = csLX711LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX711LabelEngine = new LX711LabelEngine();
                                            csLX711LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX711LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx711LabelEngine = csLX711LabelEngine.GenerateLabel();
                                            lofTasks.Add("711-LX", lx711LabelEngine);
                                        }
                                        break;
                                    case "712-LX":
                                        if (lofTasks.ContainsKey("712-LX"))
                                        {
                                            if (csLX712LabelEngine == null)
                                            {
                                                csLX712LabelEngine = new LX712LabelEngine();
                                                csLX712LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX712LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["712-LX"].IsCompleted == true &&
                                                lofTasks["712-LX"].Status != TaskStatus.Running &&
                                                lofTasks["712-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["712-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["712-LX"].Dispose();
                                                lofTasks["712-LX"] = csLX712LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX712LabelEngine = new LX712LabelEngine();
                                            csLX712LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX712LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx712LabelEngine = csLX712LabelEngine.GenerateLabel();
                                            lofTasks.Add("712-LX", lx712LabelEngine);
                                        }
                                        break;
                                    case "713-LX":
                                        if (lofTasks.ContainsKey("713-LX"))
                                        {
                                            if (csLX713LabelEngine == null)
                                            {
                                                csLX713LabelEngine = new LX713LabelEngine();
                                                csLX713LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX713LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["713-LX"].IsCompleted == true &&
                                                lofTasks["713-LX"].Status != TaskStatus.Running &&
                                                lofTasks["713-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["713-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["713-LX"].Dispose();
                                                lofTasks["713-LX"] = csLX713LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX713LabelEngine = new LX713LabelEngine();
                                            csLX713LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX713LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx713LabelEngine = csLX713LabelEngine.GenerateLabel();
                                            lofTasks.Add("713-LX", lx713LabelEngine);
                                        }
                                        break;
                                    case "717-LX":
                                        if (lofTasks.ContainsKey("717-LX"))
                                        {
                                            if (csLX717LabelEngine == null)
                                            {
                                                csLX717LabelEngine = new LX717LabelEngine();
                                                csLX717LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX717LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["717-LX"].IsCompleted == true &&
                                                lofTasks["717-LX"].Status != TaskStatus.Running &&
                                                lofTasks["717-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["717-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["717-LX"].Dispose();
                                                lofTasks["717-LX"] = csLX717LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX717LabelEngine = new LX717LabelEngine();
                                            csLX717LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX717LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx717LabelEngine = csLX717LabelEngine.GenerateLabel();
                                            lofTasks.Add("717-LX", lx717LabelEngine);
                                        }
                                        break;
                                    case "718-LX":
                                        if (lofTasks.ContainsKey("718-LX"))
                                        {
                                            if (csLX718LabelEngine == null)
                                            {
                                                csLX718LabelEngine = new LX718LabelEngine();
                                                csLX718LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX718LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["718-LX"].IsCompleted == true &&
                                                lofTasks["718-LX"].Status != TaskStatus.Running &&
                                                lofTasks["718-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["718-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["718-LX"].Dispose();
                                                lofTasks["718-LX"] = csLX718LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX718LabelEngine = new LX718LabelEngine();
                                            csLX718LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX718LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx718LabelEngine = csLX718LabelEngine.GenerateLabel();
                                            lofTasks.Add("718-LX", lx718LabelEngine);
                                        }
                                        break;
                                    case "719-LX":
                                        if (lofTasks.ContainsKey("719-LX"))
                                        {
                                            if (csLX719LabelEngine == null)
                                            {
                                                csLX719LabelEngine = new LX719LabelEngine();
                                                csLX719LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX719LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["719-LX"].IsCompleted == true &&
                                                lofTasks["719-LX"].Status != TaskStatus.Running &&
                                                lofTasks["719-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["719-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["719-LX"].Dispose();
                                                lofTasks["719-LX"] = csLX719LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX719LabelEngine = new LX719LabelEngine();
                                            csLX719LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX719LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx719LabelEngine = csLX719LabelEngine.GenerateLabel();
                                            lofTasks.Add("719-LX", lx719LabelEngine);
                                        }
                                        break;
                                    case "720-LX":
                                        if (lofTasks.ContainsKey("720-LX"))
                                        {
                                            if (csLX720LabelEngine == null)
                                            {
                                                csLX720LabelEngine = new LX720LabelEngine();
                                                csLX720LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX720LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["720-LX"].IsCompleted == true &&
                                                lofTasks["720-LX"].Status != TaskStatus.Running &&
                                                lofTasks["720-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["720-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["720-LX"].Dispose();
                                                lofTasks["720-LX"] = csLX720LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX720LabelEngine = new LX720LabelEngine();
                                            csLX720LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX720LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx720LabelEngine = csLX720LabelEngine.GenerateLabel();
                                            lofTasks.Add("720-LX", lx720LabelEngine);
                                        }
                                        break;
                                    case "721-LX":
                                        if (lofTasks.ContainsKey("721-LX"))
                                        {
                                            if (csLX721LabelEngine == null)
                                            {
                                                csLX721LabelEngine = new LX721LabelEngine();
                                                csLX721LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX721LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["721-LX"].IsCompleted == true &&
                                                lofTasks["721-LX"].Status != TaskStatus.Running &&
                                                lofTasks["721-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["721-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["721-LX"].Dispose();
                                                lofTasks["721-LX"] = csLX721LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX721LabelEngine = new LX721LabelEngine();
                                            csLX721LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX721LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx721LabelEngine = csLX721LabelEngine.GenerateLabel();
                                            lofTasks.Add("721-LX", lx721LabelEngine);
                                        }
                                        break;
                                    case "723-LX":
                                        if (lofTasks.ContainsKey("723-LX"))
                                        {
                                            if (csLX723LabelEngine == null)
                                            {
                                                csLX723LabelEngine = new LX723LabelEngine();
                                                csLX723LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX723LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["723-LX"].IsCompleted == true &&
                                                lofTasks["723-LX"].Status != TaskStatus.Running &&
                                                lofTasks["723-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["723-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["723-LX"].Dispose();
                                                lofTasks["723-LX"] = csLX723LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX723LabelEngine = new LX723LabelEngine();
                                            csLX723LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX723LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx723LabelEngine = csLX723LabelEngine.GenerateLabel();
                                            lofTasks.Add("723-LX", lx723LabelEngine);
                                        }
                                        break;
                                    case "724-LX":
                                        if (lofTasks.ContainsKey("724-LX"))
                                        {
                                            if (csLX724LabelEngine == null)
                                            {
                                                csLX724LabelEngine = new LX724LabelEngine();
                                                csLX724LabelEngine.lofFileData = new List<FileInfo>();
                                            }
                                            csLX724LabelEngine.lofFileData.Add(routingFiles[0]);
                                            if (lofTasks["724-LX"].IsCompleted == true &&
                                                lofTasks["724-LX"].Status != TaskStatus.Running &&
                                                lofTasks["724-LX"].Status != TaskStatus.WaitingToRun &&
                                                lofTasks["724-LX"].Status != TaskStatus.WaitingForActivation)
                                            {
                                                lofTasks["724-LX"].Dispose();
                                                lofTasks["724-LX"] = csLX724LabelEngine.GenerateLabel();
                                            }
                                        }
                                        else
                                        {
                                            csLX724LabelEngine = new LX724LabelEngine();
                                            csLX724LabelEngine.lofFileData = new List<FileInfo>();
                                            csLX724LabelEngine.lofFileData.Add(routingFiles[0]);
                                            Task lx724LabelEngine = csLX724LabelEngine.GenerateLabel();
                                            lofTasks.Add("724-LX", lx724LabelEngine);
                                        }
                                        break;
                                }
                            }
                        }
                        lofLines = null;
                        if(!foundMatch)
                        {
                            FileEngine csFileInputEngine = new FileEngine();
                            csFileInputEngine.MoveFileToArchive(routingFiles[0], "No Destination");
                        }
                        routingFiles.RemoveAt(0);
                    }
                    catch(Exception ex)
                    {
                        try
                        {
                            LogEngine logEngine = new LogEngine();
                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "TASK FAILED", routingFiles[0].Name + " : " + ex.ToString());
                            routingFiles.RemoveAt(0);
                        }
                        catch
                        {

                        }
                    }
                }

                if (lofTasks.Count > 0)
                {
                    LogEngine logEngine = new LogEngine();
                    List<string> lofCompletedTasks = new List<string>();
                    foreach (KeyValuePair<string, Task> task in lofTasks)
                    {
                        if (task.Value.IsCompleted == true || task.Value.Status == TaskStatus.RanToCompletion || task.Value.Status == TaskStatus.Faulted)
                        {
                            task.Value.Dispose();
                            lofCompletedTasks.Add(task.Key);

                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();

                            logEngine.WriteLog(Thread.CurrentThread.ManagedThreadId, "Tasks", "Disposed Label Task: " + task.Key);
                        }
                    }

                    foreach (string task in lofCompletedTasks)
                    {
                        if (task == "702-LX")
                        {
                            csLX702LabelEngine.Dispose();
                            csLX702LabelEngine = null;
                        }
                        else if (task == "703-LX")
                        {
                            csLX703LabelEngine.Dispose();
                            csLX703LabelEngine = null;
                        }
                        else if (task == "704-LX")
                        {
                            csLX704LabelEngine.Dispose();
                            csLX704LabelEngine = null;
                        }
                        else if (task == "706-LX")
                        {
                            csLX706LabelEngine.Dispose();
                            csLX706LabelEngine = null;
                        }
                        else if (task == "707-LX")
                        {
                            csLX707LabelEngine.Dispose();
                            csLX707LabelEngine = null;
                        }
                        else if (task == "708-LX")
                        {
                            csLX708LabelEngine.Dispose();
                            csLX708LabelEngine = null;
                        }
                        else if (task == "710-LX")
                        {
                            csLX710LabelEngine.Dispose();
                            csLX710LabelEngine = null;
                        }
                        else if (task == "711-LX")
                        {
                            csLX711LabelEngine.Dispose();
                            csLX711LabelEngine = null;
                        }
                        else if (task == "712-LX")
                        {
                            csLX712LabelEngine.Dispose();
                            csLX712LabelEngine = null;
                        }
                        else if (task == "713-LX")
                        {
                            csLX713LabelEngine.Dispose();
                            csLX713LabelEngine = null;
                        }
                        else if (task == "717-LX")
                        {
                            csLX717LabelEngine.Dispose();
                            csLX717LabelEngine = null;
                        }
                        else if (task == "718-LX")
                        {
                            csLX718LabelEngine.Dispose();
                            csLX718LabelEngine = null;
                        }
                        else if (task == "719-LX")
                        {
                            csLX719LabelEngine.Dispose();
                            csLX719LabelEngine = null;
                        }
                        else if (task == "720-LX")
                        {
                            csLX720LabelEngine.Dispose();
                            csLX720LabelEngine = null;
                        }
                        else if (task == "721-LX")
                        {
                            csLX721LabelEngine.Dispose();
                            csLX721LabelEngine = null;
                        }
                        else if (task == "723-LX")
                        {
                            csLX723LabelEngine.Dispose();
                            csLX723LabelEngine = null;
                        }
                        else if (task == "724-LX")
                        {
                            csLX724LabelEngine.Dispose();
                            csLX724LabelEngine = null;
                        }
                        lofTasks.Remove(task);

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                    }
                    lofCompletedTasks.Clear();
                }

                if (routingFiles.Count == 0 && lofTasks.Count == 0)
                {
                    break;
                }
            }
        }
    }

     public class LabelCriteriaSettings
    {
        public List<List<CriteriaSetting>> lofLabelCriteriaSettings
        {
            get
            {
                return new List<List<CriteriaSetting>>()
                {
                    LX702CriteriaSettings,
                    LX703CriteriaSettings,
                    LX704CriteriaSettings,
                    LX706CriteriaSettings,
                    LX707CriteriaSettings,
                    LX708CriteriaSettings,
                    LX710CriteriaSettings,
                    LX711CriteriaSettings,
                    LX712CriteriaSettings,
                    LX713CriteriaSettings,
                    LX717CriteriaSettings,
                    LX718CriteriaSettings,
                    LX719CriteriaSettingsForJ1,
                    LX719CriteriaSettingsForS1,
                    LX719CriteriaSettingsForBarcodeScanning2023,
                    LX720CriteriaSettings,
                    LX721CriteriaSettingsForNW,
                    LX721CriteriaSettingsForL4,
                    LX723CriteriaSettings,
                    LX724CriteriaSettings,
                };
            }
        }

        #region Label Criteria
        public List<CriteriaSetting> LX702CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Kit Production Label",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 46,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "702-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 32,
                        CriteriaRowNumber = 5,
                        RoutingLabel = "702-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX703CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Kit Production Label",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 46,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "703-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "ELCRVVW",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 33,
                        CriteriaRowNumber = 5,
                        RoutingLabel = "703-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX704CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Kit Production Label",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 46,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "704-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "ELCRVBMW",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 34,
                        CriteriaRowNumber = 5,
                        RoutingLabel = "704-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX706CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Kit Production Label",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 46,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "706-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "ELCRVTOY",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 36,
                        CriteriaRowNumber = 5,
                        RoutingLabel = "706-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX707CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Kit Production Label",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 46,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "707-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "ELCRVTOYP",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 40,
                        CriteriaRowNumber = 5,
                        RoutingLabel = "707-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX708CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Kit Production Label",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 46,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "708-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "ELCRVTOYC",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 40,
                        CriteriaRowNumber = 5,
                        RoutingLabel = "708-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX710CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "ELFUTYA - Receipts Booked",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 25,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "710-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX711CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Receipts Booked",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 15,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "711-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "V1",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 3,
                        CriteriaRowNumber = 8,
                        RoutingLabel = "711-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX712CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Receipts Booked",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 20,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "712-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX713CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Kit Production Label",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 46,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "713-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "ELCRVFORD",
                        CriteriaColumnStart = 27,
                        CriteriaColumnEnd = 37,
                        CriteriaRowNumber = 5,
                        RoutingLabel = "713-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX717CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Production - QA Approved",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 49,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "717-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX718CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Autoneum Production - QA Approved",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 57,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "718-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX719CriteriaSettingsForJ1
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Transfer Booked",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 39,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "719-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "J1",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 27,
                        CriteriaRowNumber = 8,
                        RoutingLabel = "719-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX719CriteriaSettingsForS1
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Transfer Booked",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 39,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "719-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "S1",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 27,
                        CriteriaRowNumber = 8,
                        RoutingLabel = "719-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX719CriteriaSettingsForBarcodeScanning2023
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Transfer Booked",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 39,
                        CriteriaRowNumber = 3,
                        RoutingLabel = "719-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "Barcode Scanning 2023",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 46,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "719-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX720CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "",
                        CriteriaColumnStart = 0,
                        CriteriaColumnEnd = 0,
                        CriteriaRowNumber = 0,
                        RoutingLabel = "720-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX721CriteriaSettingsForNW
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Transfer Booked",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 39,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "721-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "=NW",
                        CriteriaColumnStart = 24,
                        CriteriaColumnEnd = 27,
                        CriteriaRowNumber = 20,
                        RoutingLabel = "721-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX721CriteriaSettingsForL4
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Transfer Booked",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 39,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "721-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "=L4",
                        CriteriaColumnStart = 24,
                        CriteriaColumnEnd = 26,
                        CriteriaRowNumber = 7,
                        RoutingLabel = "721-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX723CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "Trim PE Transfer Booked",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 48,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "723-LX"
                    }
                };
            }
        }
        public List<CriteriaSetting> LX724CriteriaSettings
        {
            get
            {
                return new List<CriteriaSetting>()
                {
                    new CriteriaSetting()
                    {
                        MatchString = "V1",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 27,
                        CriteriaRowNumber = 8,
                        RoutingLabel = "724-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "TG",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 26,
                        CriteriaRowNumber = 20,
                        RoutingLabel = "724-LX"
                    },
                    new CriteriaSetting()
                    {
                        MatchString = "Transfer Booked",
                        CriteriaColumnStart = 25,
                        CriteriaColumnEnd = 39,
                        CriteriaRowNumber = 2,
                        RoutingLabel = "724-LX"
                    }
                };
            }
        }
        #endregion
    }

    public class CriteriaSetting
    {
        public string MatchString { get; set; }
        public int CriteriaColumnStart { get; set; }
        public int CriteriaColumnEnd { get; set; }
        public int CriteriaRowNumber { get; set; }
        public string RoutingLabel { get; set; }
    }
}
