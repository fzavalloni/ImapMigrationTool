using System;
using System.Windows.Forms;
using System.IO;
using ImapMigrationTool.Entities;
using System.Xml;
using System.Collections.Generic;


namespace ImapMigrationTool
{
    public static class Tools
    {
        private const string fileSettings = @".\Settings.xml";

        public static void SetRowValue(ref DataGridViewRow row, EColumns col, string obj)
        {
            try
            {
                row.Cells[col.ToString()].Value = obj;
            }
            catch { }
        }

        public static string GetPasswordOnList(string mailbox, List<SourceMailboxData> list)
        {
            SourceMailboxData mailboxData = list.Find(i => i.Mailbox == mailbox);
            return mailboxData.Password;
        }
        public static void WriteLog(string line, string targetMbx)
        {
            ConfigData data = new ConfigData();
            StreamWriter writer = null;
            try
            {
                string logPath = string.Format(string.Format("{0}{1}.txt", data.LogPath, targetMbx));

                if (Directory.Exists(data.LogPath))
                {
                    if (File.Exists(logPath))
                    {
                        writer = File.AppendText(logPath);
                    }
                    else
                    {
                        writer = new StreamWriter(logPath);
                    }

                    writer.WriteLine(line);
                }
            }
            catch (Exception erro)
            {
                throw new Exception("Failed to write in log file: " + erro.Message);
            }
            finally
            {
                if (writer != null)
                    writer.Close();
            }
        }

        public static string ReadXmlAttribute(string key)
        {
            try
            {
                using (XmlTextReader reader = new XmlTextReader(fileSettings))
                {
                    while (reader.Read())
                    {
                        if (reader.Name == key)
                        {
                            return reader.ReadElementString();
                        }
                    }
                    return "not found";
                }
            }
            catch (Exception erro)
            {
                throw new Exception("Error reading XML: " + erro.Message);
            }
        }

        public static void WriteXmlAttribute(string key, object value)
        {

            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(fileSettings);

                XmlNode root = doc.DocumentElement[key];
                root.FirstChild.InnerText = Convert.ToString(value);

                doc.Save(fileSettings);
            }
            catch (Exception erro)
            {
                throw new Exception("Error writing XML: " + erro.Message);
            }
        }

        public static void CheckImapSyncFile()
        {
            if (!File.Exists(@".\imapsync.exe"))
            {
                throw new Exception("ImapSync file is missing");
            }
        }

        public static int GetImapsyncProcessCount()
        {
            System.Diagnostics.Process[] imapsyncProcs = System.Diagnostics.Process.GetProcessesByName("imapsync");

            return imapsyncProcs.Length;
        }

        public static void KillImapsyncProcess()
        {
            System.Diagnostics.Process[] imapsyncProcs = System.Diagnostics.Process.GetProcessesByName("imapsync");

            foreach (System.Diagnostics.Process obj in imapsyncProcs)
            {
                obj.Kill();
            }
        }
    }
}
