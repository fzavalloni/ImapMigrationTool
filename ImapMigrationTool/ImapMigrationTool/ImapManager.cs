using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ActiveUp.Net.Mail;

namespace ImapMigrationTool
{
    public class ImapManager
    {
        Imap4Client client = null;
        private string adminAccount = string.Empty;
        private string adminPassword = string.Empty;

        public ImapManager(string mailServer, int port, bool ssl, string login, string password)
        {
            adminAccount = login;
            adminPassword = password;

            client = new Imap4Client();
            if (ssl)
            {
                client.ConnectSsl(mailServer, port);
            }
            else
            {
                client.Connect(mailServer, port);
            }
        }
        
        public void Connect()
        {
            //Autentica o usuário
            client.Login(adminAccount, adminPassword);                      
        }

        public void Connect(string mailboxImpersonated)
        {            
            //Concatena mailbox de admin e conta que será impersonalizada
            string aut = string.Format(@"{0}/{1}", adminAccount, mailboxImpersonated);
            
            //Autentica o usuário
            client.Login(aut, adminPassword);                                                          
        }
    }
}
