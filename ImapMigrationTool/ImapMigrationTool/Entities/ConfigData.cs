using System;

namespace ImapMigrationTool.Entities
{
    public static class ThreadData
    {
        public static int MaxThreads
        {
            get
            {
                ConfigData data = new ConfigData();
                return Convert.ToInt32(data.MaxThreads);
            }
        }
    }

    public class ConfigData
    {
        private string srcServer = Tools.ReadXmlAttribute("srcServer");

        public string SourceServer
        {
            get
            {
                return srcServer;
            }

            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The field Source server cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("srcServer", value);
                }
            }
        }

        private string srcAdminUser = Tools.ReadXmlAttribute("srcAdminUser");

        public string SourceAdminUser
        {
            get
            {
                return srcAdminUser;
            }

            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The field Source admin user cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("srcAdminUser", value);
                }
            }
        }

        public string SourceAdminPassword
        {
            get
            {
                string srcPassword = Tools.ReadXmlAttribute("srcAdminPassword");

                if (string.IsNullOrEmpty(srcPassword))
                {
                    return "";
                }

                string srcAdminPassword = Security.Decrypt(srcPassword, true);

                return srcAdminPassword;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The field Source admin password cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("srcAdminPassword", Security.Encrypt(value, true));
                }
            }
        }

        private string srcImapPort = Tools.ReadXmlAttribute("srcImapPort");

        public string SourceImapPort
        {
            get
            {
                return srcImapPort;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The field Source imap port cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("srcImapPort", value);
                }
            }
        }

        private bool srcSSL = Convert.ToBoolean(Tools.ReadXmlAttribute("srcSSL"));

        public bool SourceSSL
        {
            get
            {
                return srcSSL;
            }
            set
            {
                Tools.WriteXmlAttribute("srcSSL", Convert.ToString(value));
            }
        }

        private string tgtServer = Tools.ReadXmlAttribute("tgtServer");

        public string TargetServer
        {
            get
            {
                return tgtServer;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The field Target Server cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("tgtServer", value);
                }
            }
        }

        private string tgtAdminUser = Tools.ReadXmlAttribute("tgtAdminUser");

        public string TargetAdminUser
        {
            get
            {
                return tgtAdminUser;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The field Admin user cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("tgtAdminUser", value);
                }
            }
        }

        public string TargetAdminPassword
        {
            get
            {
                string tgtPassword = Tools.ReadXmlAttribute("tgtAdminPassword");

                if (string.IsNullOrEmpty(tgtPassword))
                {
                    return "";
                }

                string tgtAdminPassword = Security.Decrypt(tgtPassword, true);

                return tgtAdminPassword;
            }
            set
            {
                Tools.WriteXmlAttribute("tgtAdminPassword", Security.Encrypt(value, true));
            }
        }

        private string tgtImapPort = Tools.ReadXmlAttribute("tgtImapPort");

        public string TargetImapPort
        {
            get
            {
                return tgtImapPort;
            }

            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The field Target Imap port cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("tgtImapPort", value);
                }
            }
        }

        private bool tgtSSL = Convert.ToBoolean(Tools.ReadXmlAttribute("tgtSSL"));

        public bool TargetSSL
        {
            get
            {
                return tgtSSL;
            }
            set
            {
                Tools.WriteXmlAttribute("tgtSSL", value);
            }
        }

        private bool logEnabled = Convert.ToBoolean(Tools.ReadXmlAttribute("logEnabled"));

        public bool LogEnabled
        {
            get
            {
                return logEnabled;
            }
            set
            {
                Tools.WriteXmlAttribute("logEnabled", value);
            }
        }

        private string logPath = Tools.ReadXmlAttribute("logPath");

        public string LogPath
        {
            get
            {
                return logPath;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The field log path cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("logPath", value);
                }
            }
        }

        private string maxThreads = Tools.ReadXmlAttribute("maxThreads");

        public string MaxThreads
        {
            get
            {
                return maxThreads;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The field Max Threads cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("maxThreads", value);
                }
            }
        }

        private string srcServerVersion = Tools.ReadXmlAttribute("srcServerVerion");

        public string SrcServerVersion
        {
            get
            {
                return srcServerVersion;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The field Source Server Version cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("srcServerVerion", value);
                }
            }
        }

        private string tgtServerVersion = Tools.ReadXmlAttribute("tgtServerVersion");

        public string TgtServerVersion
        {
            get
            {
                return tgtServerVersion;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new Exception("The filed Target Server Version cannot be empty");
                }
                else
                {
                    Tools.WriteXmlAttribute("tgtServerVersion", value);
                }
            }
        }       

        private string ewsMessageMigration = Tools.ReadXmlAttribute("ewsMessageMigration");

        public bool EWSMessageMigration
        {
            get
            {
                return Convert.ToBoolean(ewsMessageMigration);
            }

            set
            {
                Tools.WriteXmlAttribute("ewsMessageMigration", value);
            }
        }

        private string ewsMessage = Tools.ReadXmlAttribute("ewsMessage");

        public bool EWSMessage
        {
            get
            {
                return Convert.ToBoolean(ewsMessage);
            }

            set
            {
                Tools.WriteXmlAttribute("ewsMessage", value);
            }
        }

        private string ewsContacts = Tools.ReadXmlAttribute("ewsContacts");

        public bool EWSContacts
        {
            get
            {
                return Convert.ToBoolean(ewsContacts);
            }

            set
            {
                Tools.WriteXmlAttribute("ewsContacts", value);
            }
        }

        private string ewsCalendar = Tools.ReadXmlAttribute("ewsCalendar");

        public bool EWSCalendar
        {
            get
            {
                return Convert.ToBoolean(ewsCalendar);
            }

            set
            {
                Tools.WriteXmlAttribute("ewsCalendar", value);
            }
        }

        private string ewsTasks = Tools.ReadXmlAttribute("ewsTasks");

        public bool EWSTasks
        {
            get
            {
                return Convert.ToBoolean(ewsTasks);
            }

            set
            {
                Tools.WriteXmlAttribute("ewsTasks", value);
            }
        }

        private string ewsInboxRules = Tools.ReadXmlAttribute("ewsInboxRules");

        public bool EWSInboxRules
        {
            get
            {
                return Convert.ToBoolean(ewsInboxRules);
            }

            set
            {
                Tools.WriteXmlAttribute("ewsInboxRules", value);
            }
        }

        //removeDuplicate

        private string removeDuplicate = Tools.ReadXmlAttribute("removeDuplicate");

        public bool RemoveDuplicate
        {
            get
            {
                return Convert.ToBoolean(removeDuplicate);
            }

            set
            {
                Tools.WriteXmlAttribute("removeDuplicate", value);
            }
        }

        private string migrateDefaultFolderName = Tools.ReadXmlAttribute("migrateDefaultFolderName");

        public bool MigrateDefaultFolderName
        {
            get
            {
                return Convert.ToBoolean(migrateDefaultFolderName);
            }

            set
            {
                Tools.WriteXmlAttribute("migrateDefaultFolderName", value);
            }
        }
    }
}
