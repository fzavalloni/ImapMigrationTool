using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using Microsoft.Exchange.WebServices.Data;
using System.Net.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using ImapMigrationTool.Entities;


namespace ImapMigrationTool
{
    public class FolderIdManager
    {
        public string GetFolderId(string userMail, WellKnownFolderName folderName, ExchangeVersion exVersion)
        {
            try
            {
                ConfigData data = new ConfigData();

                string adminUserName = data.TargetAdminUser;
                string adminPassword = data.TargetAdminPassword;
                string serverHost = data.TargetServer;

                string ewsUrl = string.Format("https://{0}/ews/exchange.asmx", serverHost);
                ExchangeService exService = this.GetService(adminUserName, adminPassword, userMail, ewsUrl, exVersion);
                Folder folder = Folder.Bind(exService, folderName);

                return folder.Id.UniqueId;
            }
            catch (Exception error)
            {
                throw new Exception("Error:" + error.Message);
            }
        }


        private ExchangeService GetService(string adminUserName, string adminPassword, string userMail, string ewsUrl, ExchangeVersion exVersion)
        {
            //Ignorar validação de certificado
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(OnValidationCallback);


            ExchangeService exService = new ExchangeService(exVersion);
            exService.Url = new Uri(ewsUrl);
            exService.Credentials = new NetworkCredential(adminUserName, adminPassword);

            //Impersonating user context            
            exService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userMail);

            return exService;
        }

        private static bool OnValidationCallback(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }
    }

}
