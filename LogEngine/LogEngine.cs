using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BarcodeLabelSoftware
{
    public class LogEngine
    {
        public void WriteLog(int ThreadID, string Engine, string Message)
        {
            try
            {
                DirectoryInfo logDirectory = new DirectoryInfo(Path.Combine(ConfigurationManager.AppSettings["LabelLogFolder"] + @"\"  + DateTime.Now.ToString("yyyy-MM-dd"), Engine));
                if(!logDirectory.Exists)
                {
                    logDirectory.Create();
                }
                using (StreamWriter writer = new StreamWriter(Path.Combine(logDirectory.FullName, "BLP-LOG-FILE.txt"), append: true))
                {
                    writer.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff") + ";" + ThreadID.ToString() + ";" + Engine + ";" + Message);
                }
            }
            catch
            {

            }
        }
    }
}
