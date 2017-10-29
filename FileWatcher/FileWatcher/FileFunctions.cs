using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Configuration;


namespace FileWatcher
{
    class FileFunctions
    {
        private static FileSystemWatcher fileWatcher;
        private static String logFilePath = ConfigurationSettings.AppSettings["logFilePath"];     

        public static void initialize()
        {
            // Create a file watcher that monitors the log file for changes to lastwrite date-time
            fileWatcher = new FileSystemWatcher();
            fileWatcher.Path = Path.GetDirectoryName(logFilePath);
            fileWatcher.NotifyFilter = NotifyFilters.LastWrite;
            fileWatcher.Filter = Path.GetFileName(logFilePath);
            fileWatcher.EnableRaisingEvents = true;
            fileWatcher.Changed += new FileSystemEventHandler(OnChange);
            fileWatcher.Created += new FileSystemEventHandler(OnChange);

        }

        private static void OnChange(object source, FileSystemEventArgs e)
        {
            // Read The Log File
            var logEntry = new List<string>();
            StringBuilder log = new StringBuilder();

            if (File.Exists(logFilePath))
            {
                using (Stream stream = File.Open(logFilePath, FileMode.Open, FileAccess.Read,
                        FileShare.ReadWrite))
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        var entry = reader.ReadLine();

                        while (entry != null)
                        {
                            logEntry.Add(entry);
                            entry = reader.ReadLine();
                        }
                    }
                }
                foreach (String entry in logEntry)
                    log.AppendLine(entry);
            }
            else
            {
                log.Clear();
                log.Append(String.Format("The file {0} Can Not Be Found. Please Verify the file path is correct.", logFilePath));
            }                           

            var f = Application.OpenForms.Cast<Form>().Where(x => x.Name == "Form1").FirstOrDefault();

            // Read each line and call back to the UI thread to change the tb_LogEntry RichTextBox
            f.Controls.Find("rtb_LogEntry", true).FirstOrDefault().
                Invoke((MethodInvoker)delegate {
                    RichTextBox rtb = (RichTextBox)f.Controls.Find("rtb_LogEntry", true).FirstOrDefault();
                    rtb.Clear();
                    rtb.Focus();
                    rtb.AppendText(log.ToString());                    
            });  
        }
    }
}
