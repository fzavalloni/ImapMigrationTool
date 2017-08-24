using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Net.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;

namespace ImapMigrationTool.EwsManager
{
    public class EwsMgrManagerBase
    {
        public ExchangeService GetService(string adminUserName, string adminPassword, string userMail, string ewsUrl, ExchangeVersion exVersion)
        {
            //Utiliza conta de administrador
            //Ignorar validação de certificado
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(OnValidationCallback);            

            ExchangeService exService = new ExchangeService(exVersion);
            exService.Url = new Uri(ewsUrl);
            exService.Credentials = new NetworkCredential(adminUserName, adminPassword);
            exService.Timeout = 300000; //ms
            
            //Ifazendo impersonate da conta            
            exService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userMail);

            return exService;
        }

        public ExchangeService GetService(string userMail, string password, string ewsUrl, ExchangeVersion exVersion)
        {
            //Utiliza conta sem impersonate
            //Ignorar validação de certificado
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(OnValidationCallback);

            ExchangeService exService = new ExchangeService(exVersion);
            exService.Url = new Uri(ewsUrl);
            exService.Credentials = new NetworkCredential(userMail, password);
            exService.Timeout = 300000; //ms

            return exService;
        }        

        private static bool OnValidationCallback(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors errors)
        {            
            return true;
        }
    }
}
