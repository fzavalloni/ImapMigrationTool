using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Net.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using ImapMigrationTool.Entities;
using System.Windows.Forms;

namespace ImapMigrationTool
{
    public class EWSMigrManager
    {
        bool logEnabled = false;
        string targetMailbox = string.Empty;
        int countErrors = 0;

        #region ExchangeServices
        private ExchangeService GetService(string adminUserName, string adminPassword, string userMail, string ewsUrl, ExchangeVersion exVersion)
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

        private ExchangeService GetService(string userMail, string password, string ewsUrl, ExchangeVersion exVersion)
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

        #endregion

        public void CallEWSMailboxSize(ref DataGridViewRow row, string srcUserMail, bool isAdminMigration, string password)
        {
            try
            {
                Tools.SetRowValue(ref row, EColumns.steps, "Getting size...");
                Tools.SetRowValue(ref row, EColumns.results, string.Empty);
                ConfigData data = new ConfigData();

                //Recupera versão do Exchange e converte para Enum
                ExchangeVersion exSourceVersion = (ExchangeVersion)Enum.Parse(typeof(ExchangeVersion), data.SrcServerVersion);
                ExchangeService exSrcServer = null;

                //Montando serviços
                string ewsSrcUrl = string.Format("https://{0}/ews/exchange.asmx", data.SourceServer);
                if (isAdminMigration)
                {
                    exSrcServer = this.GetService(data.SourceAdminUser, data.SourceAdminPassword, srcUserMail, ewsSrcUrl, exSourceVersion);
                }
                else
                {
                    exSrcServer = this.GetService(srcUserMail, password, ewsSrcUrl, exSourceVersion);
                }

                string result = GetMailboxSize(exSrcServer);

                Tools.SetRowValue(ref row, EColumns.MailboxSize, result);
                Tools.SetRowValue(ref row, EColumns.steps, "Finished");

            }
            catch (Exception erro)
            {
                this.WriteLog("Erro:" + erro.Message, targetMailbox);
                throw new Exception("Erro:" + erro.Message);
            }
        }



        public void CallEWSMigration(ref DataGridViewRow row, string srcUserMail, string tgtUserMail, bool isAdminMigration, string password)
        {
            try
            {
                ConfigData data = new ConfigData();

                //Seta informacoes necessarias para log das copias de
                if (data.LogEnabled)
                {
                    logEnabled = true;
                }

                targetMailbox = tgtUserMail;

                this.WriteLog("Process started: " + DateTime.Now, targetMailbox);

                //Recupera versão do Exchange e converte para Enum
                ExchangeVersion exSourceVersion = (ExchangeVersion)Enum.Parse(typeof(ExchangeVersion), data.SrcServerVersion);
                ExchangeVersion exTargetVersion = (ExchangeVersion)Enum.Parse(typeof(ExchangeVersion), data.TgtServerVersion);

                //Montando serviços
                string ewsSrcUrl = string.Format("https://{0}/ews/exchange.asmx", data.SourceServer);
                string ewsTgtUrl = string.Format("https://{0}/ews/exchange.asmx", data.TargetServer);
                ExchangeService exTgtServer = this.GetService(data.TargetAdminUser, data.TargetAdminPassword, tgtUserMail, ewsTgtUrl, exTargetVersion);
                ExchangeService exSrcServer = null;

                if (isAdminMigration)
                {
                    exSrcServer = this.GetService(data.SourceAdminUser, data.SourceAdminPassword, srcUserMail, ewsSrcUrl, exSourceVersion);
                }
                else
                {
                    exSrcServer = this.GetService(srcUserMail, password, ewsSrcUrl, exSourceVersion);
                }

                //Mostrando o tamanho do Mailbox
                Tools.SetRowValue(ref row, EColumns.steps, "0 %");
                Tools.SetRowValue(ref row, EColumns.results, "Getting Mailbox size");
                Tools.SetRowValue(ref row, EColumns.MailboxSize, this.GetMailboxSize(exSrcServer));

                //Seta linguagem das pastas do Mailbox de destino
                //Se o servidor de destino for um Exchange 2007 ele nao faz o processo porque isso nao e suportado para essa versao
                //e suportado apenas para o EWS das versoes de Exchange 2010 ou superior.
                //O problema de fazer esse configuracao é porque ela precisa ser feita via PowerShell
                //Nao encontrei nenhuma forma de fazer isso através do EWS

                if (data.MigrateDefaultFolderName)
                {
                    if (exTargetVersion != ExchangeVersion.Exchange2007_SP1)
                    {
                        this.WriteLog("Setting the mailbox language...", tgtUserMail);
                        Tools.SetRowValue(ref row, EColumns.results, "Running powershell command to set Default Folder Names...");

                        this.SetMailboxLanguage(exSrcServer, tgtUserMail);
                    }
                }

                //Fazendo Subscribe das pastas

                Tools.SetRowValue(ref row, EColumns.results, "Creating Folder Hierarchy...");

                this.SubscribeFolders(ref row, exSrcServer, exTgtServer);

                this.WriteLog("======================Folders Subscribed Successfull======================", tgtUserMail);

                if (exSourceVersion != ExchangeVersion.Exchange2007_SP1)
                {
                    //Migrando Inbox Rules
                    Tools.SetRowValue(ref row, EColumns.results, "Migrating Inbox Rules....");
                    this.MigrateInboxRules(exSrcServer, exTgtServer);

                    //Migrando Mensagens                
                    Tools.SetRowValue(ref row, EColumns.results, string.Empty);
                }
                this.WriteLog("======================Processing Itens======================", tgtUserMail);

                this.MigrateItens(ref row, exSrcServer, exTgtServer);

            }
            catch (Exception erro)
            {
                this.WriteLog("Erro:" + erro, targetMailbox);
                throw new Exception("Erro:" + erro.Message);
            }
            finally
            {
                this.WriteLog("Process ended: " + DateTime.Now, targetMailbox);
            }
        }


        private void SetMailboxLanguage(ExchangeService exSrcServer, string tgtUserMail)
        {
            Folder inboxFolder = Folder.Bind(exSrcServer, WellKnownFolderName.Inbox);
            string inboxFolderName = inboxFolder.DisplayName;

            ImapMigrationTool.RunSpace.RunSpaceManager runSpace = new ImapMigrationTool.RunSpace.RunSpaceManager();

            if (inboxFolderName.Equals("Inbox", StringComparison.CurrentCultureIgnoreCase))
            {
                //Seta OWA Mailbox para Ingles 
                runSpace.SetMailboxFolderLanguages(tgtUserMail, "en-US");
            }
            else if (inboxFolderName.Equals("Caixa de Entrada", StringComparison.CurrentCultureIgnoreCase))
            {
                //Seta OWA Mailbox para Portugues Pt-br
                runSpace.SetMailboxFolderLanguages(tgtUserMail, "pt-BR");
            }
        }

        private void MigrateItens(ref DataGridViewRow row, ExchangeService srcService, ExchangeService tgtService)
        {
            //Carrega lista de pastas que serao migradas fazendo de/para dos Ids das pastas.
            List<FolderIdentity> folderSourceList = GetFoldersIds(srcService);
            List<FolderIdentity> folderTargetList = GetFoldersIds(tgtService);
            int totalMessagesToBeMigrated = this.SumFolderCountMessages(folderSourceList);

            Tools.SetRowValue(ref row, EColumns.steps, "0 %");
            Tools.SetRowValue(ref row, EColumns.results, "Loading IDs of objects already migrated....");
            //Carrega todos os Ids dos objetos que já foram migrados
            List<ItemIdentity> migratedListIds = this.GetIdsMigratedItems(tgtService, WellKnownFolderName.MsgFolderRoot);

            int processedItemCount = 0;
            int totalCopiedFolderitemCount = 0;
            foreach (FolderIdentity folderSource in folderSourceList)
            {
                int copiedFolderitemCount = 0;

                //Recupera objetos da origem
                List<Item> contactsListSrc = this.GetFolderItems(srcService, folderSource.FolderID);

                if (contactsListSrc.Count != 0)
                {
                    string folderProgress = "Processing folder: " + folderSource.FolderPath;
                    Tools.SetRowValue(ref row, EColumns.results, folderProgress);
                    this.WriteLog("Processing folder: " + folderSource.FolderPath, targetMailbox);

                    //Recupera objetos do destino
                    //Recupera ID da mesma pasta de destino
                    FolderId tgtFolderId = this.PickCurrentObjInList(folderSource, folderTargetList);

                    foreach (Item obj in contactsListSrc)
                    {
                        processedItemCount++;

                        //int percentCompleted = ((count / totalMessagesToBeMigrated) * 100);
                        int percentCompleted = (int)Math.Round((double)(100 * processedItemCount) / totalMessagesToBeMigrated);

                        Tools.SetRowValue(ref row, EColumns.steps, string.Format("{0} %", percentCompleted));
                        Tools.SetRowValue(ref row, EColumns.results, string.Format("{0} - Item {1}", folderProgress, processedItemCount));
                        //Caso o ID não exista no destino a cópia do objeto é feita
                        string id = obj.Id.UniqueId;


                        //Trata Appointments que tiveram alguma alteracao
                        if (obj is Appointment)
                        {
                            foreach (ItemIdentity i in migratedListIds)
                            {
                                //Verifica se o objeto já foi migrado
                                if (i.AppOldId == id)
                                {
                                    Appointment oldApp = null;
                                    //Se o appointment antigo for deletado ele nao gera erro.
                                    try
                                    {
                                        //Faz o bind do antigo
                                        oldApp = Appointment.Bind(srcService, id);
                                    }
                                    catch { }


                                    //Faz o Bind no migrado
                                    Appointment newApp = Appointment.Bind(tgtService, i.AppNewId);

                                    //Em vez de ficar comparando atributo por atributo ele deleta o objeto migrado
                                    //e sempre irá prevalecer os objetos das origem
                                    //Tambem coloca para ele não enviar uma mensagem de Reunião cancelada

                                    newApp.Delete(DeleteMode.HardDelete, SendCancellationsMode.SendToNone);

                                    if (oldApp != null)
                                    {
                                        this.SaveCalendar(tgtService, oldApp, tgtFolderId);
                                        copiedFolderitemCount++;
                                    }
                                }
                            }
                        }



                        if (!migratedListIds.Exists(i => i.AppOldId == id))
                        {
                            if (obj is Contact)
                            {
                                Contact contact = obj as Contact;
                                this.SaveContact(tgtService, contact, tgtFolderId);
                                copiedFolderitemCount++;
                            }

                            if (obj is Appointment)
                            {
                                Appointment appt = obj as Appointment;
                                this.SaveCalendar(tgtService, appt, tgtFolderId);
                                copiedFolderitemCount++;
                            }

                            if (obj is Task)
                            {
                                Task task = obj as Task;
                                this.SaveTask(tgtService, task, tgtFolderId);
                                copiedFolderitemCount++;
                            }

                            if (obj is EmailMessage)
                            {
                                EmailMessage msg = obj as EmailMessage;
                                this.SaveMessage(tgtService, msg, tgtFolderId);
                                copiedFolderitemCount++;
                            }
                        }
                    }
                }
                //Loga a quantidade de itens copiados dessa pasta
                this.WriteLog(string.Format("Folder Itens: {0} itens", folderSource.CountItems), targetMailbox);
                this.WriteLog(string.Format("Copied Itens: {0} itens", copiedFolderitemCount), targetMailbox);

                //Soma todos os itens que são copiados para mostrar o resultado no log
                totalCopiedFolderitemCount = totalCopiedFolderitemCount + copiedFolderitemCount;
                copiedFolderitemCount++;
            }

            this.WriteLog("============================", targetMailbox);
            this.WriteLog(string.Format("Processed Itens: {0} itens", processedItemCount), targetMailbox);
            this.WriteLog(string.Format("Copied Itens: {0} itens", totalCopiedFolderitemCount), targetMailbox);
            this.WriteLog(string.Format("Itens NOT migrated from Deleted Itens Folder: {0}", this.GetDeletedItensCount(srcService)), targetMailbox);
            Tools.SetRowValue(ref row, EColumns.steps, "Errors: " + countErrors);
            this.WriteLog("Errors: " + countErrors, targetMailbox);
        }

        private void MigrateInboxRules(ExchangeService srcService, ExchangeService tgtService)
        {
            try
            {
                RuleCollection ruleCollectionSource = srcService.GetInboxRules();
                RuleCollection ruleCollectionTarget = tgtService.GetInboxRules();

                if (ruleCollectionSource.Count != 0)
                {
                    foreach (Rule ruleSource in ruleCollectionSource)
                    {
                        bool isruleAlreadyCreated = this.IsRuleAlreadyCreated(ruleSource.DisplayName, ruleCollectionTarget);
                        if (!isruleAlreadyCreated)
                        {
                            CreateRuleOperation createRule = new CreateRuleOperation(this.CopyRule(ruleSource, srcService, tgtService));

                            tgtService.UpdateInboxRules(new RuleOperation[] { createRule }, true);
                        }
                    }
                }
            }

            catch (Exception erro)
            {
                this.WriteLog("Error creating Inbox Rules: " + erro.Message, targetMailbox);
            }
        }

        private bool IsRuleAlreadyCreated(string ruleDisplayName, RuleCollection ruleCollection)
        {
            bool iscreated = false;

            if (ruleCollection.Count != 0)
            {
                foreach (Rule rule in ruleCollection)
                {
                    if (ruleDisplayName == rule.DisplayName)
                        iscreated = true;
                }
            }
            return iscreated;
        }

        //Este metodo converte o FolderId da origem para o ID do destino

        private FolderId ConvertFolderId(ExchangeService srcService, ExchangeService tgtService, FolderId fId)
        {
            FolderId idConverted = null;
            string folderPath = string.Empty;
            //Recupera a lista das pastas e seus Ids e caminhos
            List<FolderIdentity> srcFolders = this.GetFoldersIds(srcService);
            List<FolderIdentity> tgtFolders = this.GetFoldersIds(tgtService);

            //Pego o caminho fisico da pasta            

            foreach (FolderIdentity srcFolder in srcFolders)
            {
                if (string.Equals(srcFolder.FolderID.ToString(), fId.ToString(), StringComparison.CurrentCultureIgnoreCase))
                {
                    folderPath = srcFolder.FolderPath;
                }
            }
            //No destino recupero o id baseado no path fisico
            foreach (FolderIdentity tgtFolder in tgtFolders)
            {
                if (string.Equals(folderPath, tgtFolder.FolderPath, StringComparison.CurrentCultureIgnoreCase))
                {
                    idConverted = tgtFolder.FolderID;
                }
            }

            return idConverted;
        }

        private Rule CopyRule(Rule oldRule, ExchangeService srcService, ExchangeService tgtService)
        {
            Rule newRule = null;
            try
            {
                newRule = new Rule();
                if (oldRule.Actions.AssignCategories.Count != 0)
                {
                    foreach (string obj in oldRule.Actions.AssignCategories)
                    {
                        newRule.Actions.AssignCategories.Add(obj);
                    }
                }

                //Todas as regras de move e copy para uma pasta se baseiam no ID da pasta
                //Sendo assim, foi preciso fazer um metodo que pegasse o ID da mesma pasta no destino
                if (oldRule.Actions.CopyToFolder != null)
                {
                    newRule.Actions.CopyToFolder = this.ConvertFolderId(srcService, tgtService, oldRule.Actions.CopyToFolder);
                }

                if (oldRule.Actions.MoveToFolder != null)
                {
                    newRule.Actions.MoveToFolder = this.ConvertFolderId(srcService, tgtService, oldRule.Actions.MoveToFolder);
                }

                newRule.Actions.Delete = oldRule.Actions.Delete;

                if (oldRule.Actions.ForwardAsAttachmentToRecipients.Count != 0)
                {
                    foreach (EmailAddress s in oldRule.Actions.ForwardAsAttachmentToRecipients)
                    {
                        //Se o endereço for interno ele retorna um legacyExchangeDN, ai temos que tratar isso
                        //porque ele não deixa adicionar legacyDN, assim temos que converter 
                        if (s.Address.Contains("/cn"))
                        {
                            string email = ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
                            if (email != null)
                            {
                                newRule.Actions.ForwardAsAttachmentToRecipients.Add(email);
                            }
                        }
                        else
                        {
                            newRule.Actions.ForwardAsAttachmentToRecipients.Add(s);
                        }
                    }
                }

                if (oldRule.Actions.ForwardToRecipients.Count != 0)
                {
                    foreach (EmailAddress s in oldRule.Actions.ForwardToRecipients)
                    {
                        //Se o endereço for interno ele retorna um legacyExchangeDN, ai temos que tratar isso
                        //porque ele não deixa adicionar legacyDN, assim temos que converter 
                        if (s.Address.Contains("/cn"))
                        {
                            string email = ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
                            if (email != null)
                            {
                                newRule.Actions.ForwardToRecipients.Add(email);
                            }
                        }
                        else
                        {
                            newRule.Actions.ForwardToRecipients.Add(s);
                        }
                    }
                }
                newRule.Actions.MarkAsRead = oldRule.Actions.MarkAsRead;
                newRule.Actions.MarkImportance = oldRule.Actions.MarkImportance;
                newRule.Actions.PermanentDelete = oldRule.Actions.PermanentDelete;
                if (oldRule.Actions.RedirectToRecipients.Count != 0)
                {
                    foreach (EmailAddress s in oldRule.Actions.RedirectToRecipients)
                    {
                        //Se o endereço for interno ele retorna um legacyExchangeDN, ai temos que tratar isso
                        //porque ele não deixa adicionar legacyDN, assim temos que converter 
                        if (s.Address.Contains("/cn"))
                        {
                            string email = ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
                            if (email != null)
                            {
                                newRule.Actions.RedirectToRecipients.Add(email);
                            }
                        }
                        else
                        {
                            newRule.Actions.RedirectToRecipients.Add(s);
                        }
                    }
                }
                if (oldRule.Actions.SendSMSAlertToRecipients.Count != 0)
                {
                    foreach (MobilePhone mobile in oldRule.Actions.SendSMSAlertToRecipients)
                    {
                        newRule.Actions.SendSMSAlertToRecipients.Add(mobile);
                    }
                }
                oldRule.Actions.ServerReplyWithMessage = newRule.Actions.ServerReplyWithMessage;
                oldRule.Actions.StopProcessingRules = newRule.Actions.StopProcessingRules;

                if (oldRule.Conditions.Categories.Count != 0)
                {
                    foreach (string s in oldRule.Conditions.Categories)
                    {
                        newRule.Conditions.Categories.Add(s);
                    }
                }
                newRule.DisplayName = oldRule.DisplayName;

                if (oldRule.Conditions.ContainsBodyStrings != null)
                {
                    foreach (string s in oldRule.Conditions.ContainsBodyStrings)
                    {
                        newRule.Conditions.ContainsBodyStrings.Add(s);
                    }
                }

                if (oldRule.Conditions.ContainsHeaderStrings != null)
                {
                    foreach (string s in oldRule.Conditions.ContainsHeaderStrings)
                    {
                        newRule.Conditions.ContainsHeaderStrings.Add(s);
                    }
                }

                if (oldRule.Conditions.FromAddresses != null)
                {
                    foreach (EmailAddress s in oldRule.Conditions.FromAddresses)
                    {
                        //Se o endereço for interno ele retorna um legacyExchangeDN, ai temos que tratar isso
                        //porque ele não deixa adicionar legacyDN, assim temos que converter 
                        if (s.Address.Contains("/cn"))
                        {
                            string email = ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
                            if (email != null)
                            {
                                newRule.Conditions.FromAddresses.Add(email);
                            }
                        }
                        else
                        {
                            newRule.Conditions.FromAddresses.Add(s);
                        }
                    }
                }

                if (oldRule.Conditions.ContainsRecipientStrings != null)
                {
                    foreach (string s in oldRule.Conditions.ContainsRecipientStrings)
                    {
                        newRule.Conditions.ContainsRecipientStrings.Add(s);
                    }
                }

                if (oldRule.Conditions.ContainsSenderStrings != null)
                {
                    foreach (string s in oldRule.Conditions.ContainsSenderStrings)
                    {
                        newRule.Conditions.ContainsSenderStrings.Add(s);
                    }
                }

                if (oldRule.Conditions.ContainsSubjectOrBodyStrings != null)
                {
                    foreach (string s in oldRule.Conditions.ContainsSubjectOrBodyStrings)
                    {
                        newRule.Conditions.ContainsSubjectOrBodyStrings.Add(s);
                    }
                }

                if (oldRule.Conditions.ContainsSubjectStrings != null)
                {
                    foreach (string s in oldRule.Conditions.ContainsSubjectStrings)
                    {
                        newRule.Conditions.ContainsSubjectStrings.Add(s);
                    }
                }

                newRule.Conditions.FlaggedForAction = oldRule.Conditions.FlaggedForAction;

                if (oldRule.Conditions.FromConnectedAccounts != null)
                {
                    foreach (string s in oldRule.Conditions.FromConnectedAccounts)
                    {
                        newRule.Conditions.FromConnectedAccounts.Add(s);
                    }
                }

                newRule.Conditions.Importance = oldRule.Conditions.Importance;

                if (oldRule.Exceptions.Categories != null)
                {
                    foreach (string s in oldRule.Exceptions.Categories)
                    {
                        newRule.Exceptions.Categories.Add(s);
                    }
                }

                if (oldRule.Exceptions.ContainsBodyStrings != null)
                {
                    foreach (string s in oldRule.Exceptions.ContainsBodyStrings)
                    {
                        newRule.Exceptions.ContainsBodyStrings.Add(s);
                    }
                }
                if (oldRule.Exceptions.ContainsHeaderStrings != null)
                {
                    foreach (string s in oldRule.Exceptions.ContainsHeaderStrings)
                    {
                        newRule.Exceptions.ContainsHeaderStrings.Add(s);
                    }
                }
                if (oldRule.Exceptions.ContainsRecipientStrings != null)
                {
                    foreach (string s in oldRule.Exceptions.ContainsRecipientStrings)
                    {
                        newRule.Exceptions.ContainsRecipientStrings.Add(s);
                    }
                }
                if (oldRule.Exceptions.ContainsRecipientStrings != null)
                {
                    foreach (string s in oldRule.Exceptions.ContainsRecipientStrings)
                    {
                        newRule.Exceptions.ContainsRecipientStrings.Add(s);
                    }
                }
                if (oldRule.Exceptions.ContainsSenderStrings != null)
                {
                    foreach (string s in oldRule.Exceptions.ContainsSenderStrings)
                    {
                        newRule.Exceptions.ContainsSenderStrings.Add(s);
                    }
                }
                if (oldRule.Exceptions.ContainsSenderStrings != null)
                {
                    foreach (string s in oldRule.Exceptions.ContainsSenderStrings)
                    {
                        newRule.Exceptions.ContainsSenderStrings.Add(s);
                    }
                }
                if (oldRule.Exceptions.ContainsSubjectOrBodyStrings != null)
                {
                    foreach (string s in oldRule.Exceptions.ContainsSubjectOrBodyStrings)
                    {
                        newRule.Exceptions.ContainsSubjectOrBodyStrings.Add(s);
                    }
                }
                if (oldRule.Exceptions.ContainsSubjectStrings != null)
                {
                    foreach (string s in oldRule.Exceptions.ContainsSubjectStrings)
                    {
                        newRule.Exceptions.ContainsSubjectStrings.Add(s);
                    }
                }

                newRule.Exceptions.FlaggedForAction = oldRule.Exceptions.FlaggedForAction;

                if (oldRule.Exceptions.FromAddresses != null)
                {
                    foreach (EmailAddress s in oldRule.Exceptions.FromAddresses)
                    {
                        //Se o endereço for interno ele retorna um legacyExchangeDN, ai temos que tratar isso
                        //porque ele não deixa adicionar legacyDN, assim temos que converter 
                        if (s.Address.Contains("/cn"))
                        {
                            string email = ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
                            if (email != null)
                            {
                                newRule.Exceptions.FromAddresses.Add(email);
                            }
                        }
                        else
                        {
                            newRule.Exceptions.FromAddresses.Add(s);
                        }
                    }
                }
                if (oldRule.Exceptions.FromConnectedAccounts != null)
                {
                    foreach (string s in oldRule.Exceptions.FromConnectedAccounts)
                    {
                        newRule.Exceptions.FromConnectedAccounts.Add(s);
                    }
                }
                newRule.Exceptions.Importance = oldRule.Exceptions.Importance;
                newRule.Exceptions.IsApprovalRequest = oldRule.Exceptions.IsApprovalRequest;
                newRule.Exceptions.IsAutomaticForward = oldRule.Exceptions.IsAutomaticForward;
                newRule.Exceptions.IsAutomaticReply = oldRule.Exceptions.IsAutomaticReply;
                newRule.Exceptions.IsEncrypted = oldRule.Exceptions.IsEncrypted;
                newRule.Exceptions.IsMeetingRequest = oldRule.Exceptions.IsMeetingRequest;
                newRule.Exceptions.IsMeetingResponse = oldRule.Exceptions.IsMeetingResponse;
                newRule.Exceptions.IsNonDeliveryReport = oldRule.Exceptions.IsNonDeliveryReport;
                newRule.Exceptions.IsPermissionControlled = oldRule.Exceptions.IsPermissionControlled;
                newRule.Exceptions.IsReadReceipt = oldRule.Exceptions.IsReadReceipt;
                if (oldRule.Exceptions.MessageClassifications != null)
                {
                    foreach (string s in oldRule.Exceptions.MessageClassifications)
                    {
                        newRule.Exceptions.MessageClassifications.Add(s);
                    }
                }
                newRule.Exceptions.NotSentToMe = oldRule.Exceptions.NotSentToMe;
                newRule.Exceptions.Sensitivity = oldRule.Exceptions.Sensitivity;
                newRule.Exceptions.SentCcMe = oldRule.Exceptions.SentCcMe;
                newRule.Exceptions.SentOnlyToMe = oldRule.Exceptions.SentOnlyToMe;
                if (oldRule.Exceptions.SentToAddresses != null)
                {
                    foreach (EmailAddress s in oldRule.Exceptions.SentToAddresses)
                    {
                        //Se o endereço for interno ele retorna um legacyExchangeDN, ai temos que tratar isso
                        //porque ele não deixa adicionar legacyDN, assim temos que converter 
                        if (s.Address.Contains("/cn"))
                        {
                            string email = ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
                            if (email != null)
                            {
                                newRule.Exceptions.SentToAddresses.Add(email);
                            }
                        }
                        else
                        {
                            newRule.Exceptions.SentToAddresses.Add(s);
                        }
                    }
                }

                newRule.Exceptions.SentToMe = oldRule.Exceptions.SentToMe;
                newRule.Exceptions.SentToOrCcMe = oldRule.Exceptions.SentToOrCcMe;
                newRule.Exceptions.WithinDateRange.Start = oldRule.Exceptions.WithinDateRange.Start;
                newRule.Exceptions.WithinDateRange.End = oldRule.Exceptions.WithinDateRange.End;
                newRule.IsEnabled = oldRule.IsEnabled;
                newRule.Priority = oldRule.Priority;
            }
            catch (Exception erro)
            {
                this.WriteLog("Error to copy InboxRule: " + erro.Message, targetMailbox);
            }
            return newRule;
        }

        //Faz uma busca no Exchange pelo legacyDN e retorna o emailAddress
        private string ConvertLegacyExchangeDNToSmtpAddress(EmailAddress email, ExchangeService svc)
        {
            string smtp = null;
            string legacyDN = email.Address;
            NameResolutionCollection col = null;

            col = svc.ResolveName(email.Name, ResolveNameSearchLocation.DirectoryOnly, true);
            if (col.Count != 0)
            {
                smtp = col[0].Mailbox.Address;
            }
            return smtp;
        }


        private List<FolderIdentity> RemoveFoldersAlreadyExists(List<FolderIdentity> folderListSourceSrv, List<FolderIdentity> folderListTargetSrv)
        {
            List<FolderIdentity> cleanedList = new List<FolderIdentity>();

            foreach (FolderIdentity folderSource in folderListSourceSrv)
            {
                if (!folderListTargetSrv.Exists(new Predicate<FolderIdentity>(i => i.FolderPath == folderSource.FolderPath)))
                {
                    cleanedList.Add(folderSource);
                }
            }

            return cleanedList;
        }


        private void SubscribeFolders(ref DataGridViewRow row, ExchangeService exSrcServer, ExchangeService exTgtServer)
        {
            //Recupera a lista de todas as pasta na origem e destino
            List<FolderIdentity> folderListSourceSrv = GetFolderHierarchy(exSrcServer);
            List<FolderIdentity> folderListTargetSrv = GetFolderHierarchy(exTgtServer);
            //Gera uma nova lista apenas com as pastas que precisam ser criadas
            List<FolderIdentity> diffList = RemoveFoldersAlreadyExists(folderListSourceSrv, folderListTargetSrv);

            FolderId lastFolderId = null;

            foreach (FolderIdentity path in diffList)
            {
                Tools.SetRowValue(ref row, EColumns.results, "Subscribing Folder: " + path.FolderPath);

                if (path.FolderPath.Contains("\\"))
                {
                    string[] folders = path.FolderPath.Split('\\');

                    for (int i = 0; i < folders.Length; i++)
                    {

                        if (i == 0)
                        {
                            //Cria a pasta na raiz
                            CreateNewFolder(exTgtServer, folders[i], WellKnownFolderName.MsgFolderRoot, path.FolderType);
                            lastFolderId = GetFolderId(exTgtServer, WellKnownFolderName.MsgFolderRoot, folders[i]);
                        }
                        else
                        {
                            //Cria as pastas filhas 
                            CreateNewFolder(exTgtServer, folders[i], lastFolderId, path.FolderType);
                            lastFolderId = GetFolderId(exTgtServer, lastFolderId, folders[i]);
                        }
                    }
                }
                else
                {
                    //Create Folder no Root da pasta
                    CreateNewFolder(exTgtServer, path.FolderPath, WellKnownFolderName.MsgFolderRoot, path.FolderType);
                }
            }
        }


        private FolderId GetFolderId(ExchangeService svc, WellKnownFolderName searchRootFolder, string folderName)
        {
            FolderId idFolder = null;
            Folder folder = Folder.Bind(svc, searchRootFolder);
            folder.Load();

            foreach (Folder obj in folder.FindFolders(new FolderView(int.MaxValue)))
            {
                if (obj.DisplayName == folderName)
                {
                    idFolder = obj.Id;
                }
            }

            return idFolder;
        }

        private FolderId GetFolderId(ExchangeService svc, FolderId searchRootFolderId, string folderName)
        {
            FolderId idFolder = null;
            Folder folder = Folder.Bind(svc, searchRootFolderId);
            folder.Load();
            foreach (Folder obj in folder.FindFolders(new FolderView(int.MaxValue)))
            {
                if (obj.DisplayName == folderName)
                {
                    idFolder = obj.Id;
                }
            }

            return idFolder;
        }


        private void CreateNewFolder(ExchangeService svc, string folderName, FolderId folderId, string folderClass)
        {
            //Verifica o tipo de pasta e cria de acordo com o seu tipo

            switch (folderClass)
            {
                case "IPF.Note":
                    {
                        CreateMessageFolder(svc, folderName, folderId);
                        break;
                    }

                case "IPF.Appointment":
                    {
                        CreateCalendarFolder(svc, folderName, folderId);
                        break;
                    }

                case "IPF.Contact":
                    {
                        CreateContactsFolder(svc, folderName, folderId);
                        break;
                    }
                case "IPF.Task":
                    {
                        CreateTasksFolder(svc, folderName, folderId);
                        break;
                    }
                case "IPF.StickyNote":
                    {
                        CreateNotesFolder(svc, folderName, folderId, "IPF.StickyNote");
                        break;
                    }
                default:
                    {
                        //Se não for de nenhum tipo ele cria uma pasta do tipo Ipn.Note
                        CreateMessageFolder(svc, folderName, folderId);
                        break;
                    }
            }
        }

        private void CreateNewFolder(ExchangeService svc, string folderName, WellKnownFolderName folderWellKnow, string folderClass)
        {
            switch (folderClass)
            {
                case "IPF.Note":
                    {
                        CreateMessageFolder(svc, folderName, folderWellKnow);
                        break;
                    }

                case "IPF.Appointment":
                    {
                        CreateCalendarFolder(svc, folderName, folderWellKnow);
                        break;
                    }

                case "IPF.Contact":
                    {
                        CreateContactsFolder(svc, folderName, folderWellKnow);
                        break;
                    }
                case "IPF.Task":
                    {
                        CreateTasksFolder(svc, folderName, folderWellKnow);
                        break;
                    }
                case "IPF.StickyNote":
                    {
                        CreateNotesFolder(svc, folderName, folderWellKnow, "IPF.StickyNote");
                        break;
                    }
                default:
                    {
                        //Se não for de nenhum tipo, cria uma pasta 
                        CreateMessageFolder(svc, folderName, folderWellKnow);
                        break;
                    }

            }
        }

        #region Creating Folder Methods

        private void CreateMessageFolder(ExchangeService svc, string folderName, WellKnownFolderName folderWellKnow)
        {
            Folder folder = new Folder(svc);

            folder.DisplayName = folderName;

            try
            {
                folder.Save(folderWellKnow);
            }
            catch (Exception erro)
            {
                if (!erro.Message.Equals("A folder with the specified name already exists.", StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new Exception(erro.Message);
                }
            }
        }

        private void CreateMessageFolder(ExchangeService svc, string folderName, FolderId folderId)
        {
            Folder folder = new Folder(svc);

            folder.DisplayName = folderName;

            try
            {
                folder.Save(folderId);
            }
            catch (Exception erro)
            {
                if (!erro.Message.Equals("A folder with the specified name already exists.", StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new Exception(erro.Message);
                }
            }
        }

        private void CreateNotesFolder(ExchangeService svc, string folderName, WellKnownFolderName folderWellKnow, string folderClass)
        {
            Folder folder = new Folder(svc);
            folder.DisplayName = folderName;
            folder.FolderClass = folderClass;

            try
            {
                folder.Save(folderWellKnow);
            }
            catch (Exception erro)
            {
                if (!erro.Message.Equals("A folder with the specified name already exists.", StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new Exception(erro.Message);
                }
            }
        }

        private void CreateNotesFolder(ExchangeService svc, string folderName, FolderId folderId, string folderClass)
        {
            Folder folder = new Folder(svc);
            folder.DisplayName = folderName;
            folder.FolderClass = folderClass;

            try
            {
                folder.Save(folderId);
            }
            catch (Exception erro)
            {
                if (!erro.Message.Equals("A folder with the specified name already exists.", StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new Exception(erro.Message);
                }
            }
        }


        private void CreateCalendarFolder(ExchangeService svc, string folderName, WellKnownFolderName folderWellKnow)
        {
            CalendarFolder folder = new CalendarFolder(svc);

            folder.DisplayName = folderName;

            try
            {
                folder.Save(folderWellKnow);
            }
            catch (Exception erro)
            {
                if (!erro.Message.Equals("A folder with the specified name already exists.", StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new Exception(erro.Message);
                }
            }
        }

        private void CreateCalendarFolder(ExchangeService svc, string folderName, FolderId folderId)
        {
            CalendarFolder folder = new CalendarFolder(svc);
            folder.DisplayName = folderName;

            try
            {
                folder.Save(folderId);
            }
            catch (Exception erro)
            {
                if (!erro.Message.Equals("A folder with the specified name already exists.", StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new Exception(erro.Message);
                }
            }
        }

        private void CreateTasksFolder(ExchangeService svc, string folderName, WellKnownFolderName folderWellKnow)
        {
            TasksFolder folder = new TasksFolder(svc);

            folder.DisplayName = folderName;

            try
            {
                folder.Save(folderWellKnow);
            }
            catch (Exception erro)
            {
                if (!erro.Message.Equals("A folder with the specified name already exists.", StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new Exception(erro.Message);
                }
            }
        }

        private void CreateTasksFolder(ExchangeService svc, string folderName, FolderId folderId)
        {
            TasksFolder folder = new TasksFolder(svc);

            folder.DisplayName = folderName;

            try
            {
                folder.Save(folderId);
            }
            catch (Exception erro)
            {
                if (!erro.Message.Equals("A folder with the specified name already exists.", StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new Exception(erro.Message);
                }
            }
        }

        private void CreateContactsFolder(ExchangeService svc, string folderName, WellKnownFolderName folderWellKnow)
        {
            ContactsFolder folder = new ContactsFolder(svc);

            folder.DisplayName = folderName;

            try
            {
                folder.Save(folderWellKnow);
            }
            catch (Exception erro)
            {
                if (!erro.Message.Equals("A folder with the specified name already exists.", StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new Exception(erro.Message);
                }
            }
        }

        private void CreateContactsFolder(ExchangeService svc, string folderName, FolderId folderId)
        {
            ContactsFolder folder = new ContactsFolder(svc);

            folder.DisplayName = folderName;

            try
            {
                folder.Save(folderId);
            }
            catch (Exception erro)
            {
                if (!erro.Message.Equals("A folder with the specified name already exists.", StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new Exception(erro.Message);
                }
            }
        }

        #endregion

        private string GetMailboxSize(ExchangeService svc)
        {
            FolderView view = new FolderView(int.MaxValue);
            view.Traversal = FolderTraversal.Deep;

            PropertySet ps = new PropertySet(BasePropertySet.FirstClassProperties);
            ExtendedPropertyDefinition prMessageSizeExt = new ExtendedPropertyDefinition(3592, MapiPropertyType.Long);
            ExtendedPropertyDefinition prDeletedMsgSizeExt = new ExtendedPropertyDefinition(26267, MapiPropertyType.Long);
            ExtendedPropertyDefinition prDeletedMsgCount = new ExtendedPropertyDefinition(26176, MapiPropertyType.Integer);

            ps.Add(prMessageSizeExt);
            ps.Add(prDeletedMsgSizeExt);
            ps.Add(prDeletedMsgCount);

            view.PropertySet = ps;

            FindFoldersResults foldersResults = svc.FindFolders(WellKnownFolderName.MsgFolderRoot, view);

            Int64 totalItemCount = 0;
            Int64 totalItemSize = 0;

            foreach (Folder folder in foldersResults)
            {
                totalItemCount = totalItemCount + folder.TotalCount;
                Int64 folderSize;
                if (folder.TryGetProperty(prMessageSizeExt, out folderSize))
                {
                    totalItemSize = totalItemSize + Convert.ToInt64(folderSize);
                }
                Int64 deletedItemFolderSize;
                if (folder.TryGetProperty(prDeletedMsgSizeExt, out deletedItemFolderSize))
                {
                    totalItemSize = totalItemSize + Convert.ToInt64(deletedItemFolderSize);
                }
            }

            //Converte o tamanho de byte para MB/ O decimal.Round deixa apenas 2 casas decimais
            decimal totalItemSizeInMB = decimal.Round(Convert.ToDecimal(((totalItemSize / 1024f) / 1024f)), 2);

            return string.Format("Size: {0} MB | Itens: {1}", totalItemSizeInMB, totalItemCount);
        }

        private List<FolderIdentity> GetFolderHierarchy(ExchangeService svc)
        {
            List<FolderIdentity> pathList = new List<FolderIdentity>();
            //Adiciona propriedade de FolderPath Ex: \\Inbox\\Folder1\\Folder2
            ExtendedPropertyDefinition folderPath = new ExtendedPropertyDefinition(26293, MapiPropertyType.String);
            FolderView view = new FolderView(int.MaxValue);
            view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties) { folderPath };
            view.PropertySet.Add(FolderSchema.ChildFolderCount);
            view.PropertySet.Add(FolderSchema.TotalCount);
            view.PropertySet.Add(FolderSchema.ParentFolderId);
            view.PropertySet.Add(FolderSchema.DisplayName);
            //Faz busca nas pastas filhas
            view.Traversal = FolderTraversal.Deep;

            foreach (Folder obj in svc.FindFolders(WellKnownFolderName.MsgFolderRoot, view))
            {
                Object fpPath = null;
                obj.TryGetProperty(folderPath, out fpPath);

                //Remove as \\ do comeco do path e adiciona
                string path = fpPath.ToString().Remove(0, 1);
                pathList.Add(new FolderIdentity(path, obj.FolderClass, obj.TotalCount));
            }

            return pathList;
        }


        private List<FolderIdentity> GetFoldersIds(ExchangeService svc)
        {
            List<FolderIdentity> pathList = new List<FolderIdentity>();
            //Adiciona propriedade de FolderPath Ex: \\Inbox\\Folder1\\Folder2
            ExtendedPropertyDefinition folderPath = new ExtendedPropertyDefinition(26293, MapiPropertyType.String);
            FolderView view = new FolderView(int.MaxValue);
            view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties) { folderPath };
            view.PropertySet.Add(FolderSchema.ChildFolderCount);
            view.PropertySet.Add(FolderSchema.TotalCount);
            view.PropertySet.Add(FolderSchema.ParentFolderId);
            view.PropertySet.Add(FolderSchema.DisplayName);
            //Faz busca das pastas filhas
            view.Traversal = FolderTraversal.Deep;

            //Nao devemos adicionar na busca a pasta de Itens Excluidos
            Folder folderDeletedItens = Folder.Bind(svc, WellKnownFolderName.DeletedItems);

            foreach (Folder obj in svc.FindFolders(WellKnownFolderName.MsgFolderRoot, view))
            {
                //Nao adiciona a pasta Deleted Itens na busca
                if (!folderDeletedItens.Id.UniqueId.Equals(obj.Id.UniqueId))
                {
                    Object fpPath = null;
                    obj.TryGetProperty(folderPath, out fpPath);

                    //Remove as \\ do comeco do path e adiciona
                    string path = fpPath.ToString().Remove(0, 1);

                    pathList.Add(new FolderIdentity(path, obj.FolderClass, obj.Id, obj.TotalCount));
                }
            }
            return pathList;
        }


        private void WriteLog(string message, string targetMailbox)
        {
            if (logEnabled)
            {
                Tools.WriteLog(message, targetMailbox);
            }
        }

        private FolderId PickCurrentObjInList(FolderIdentity obj, List<FolderIdentity> list)
        {
            FolderId id = null;

            foreach (FolderIdentity item in list)
            {
                if (obj.FolderPath == item.FolderPath)
                {
                    id = item.FolderID;
                }
            }
            return id;
        }

        private int SumFolderCountMessages(List<FolderIdentity> list)
        {
            int sum = 0;

            foreach (FolderIdentity item in list)
            {
                sum = sum + item.CountItems;
            }

            return sum;
        }

        private int GetDeletedItensCount(ExchangeService svc)
        {
            Folder folder = Folder.Bind(svc, WellKnownFolderName.DeletedItems);
            return folder.TotalCount;
        }

        private TimeSpan DateTimeToTimeSpan(DateTime? ts)
        {
            if (!ts.HasValue)
            {
                return TimeSpan.Zero;
            }
            else
            {
                return new TimeSpan(0, ts.Value.Hour, ts.Value.Minute, ts.Value.Second, ts.Value.Millisecond);
            }
        }

        private void SaveMessage(ExchangeService srv, EmailMessage msg, FolderId folderId)
        {
            EmailMessage emailMsg = new EmailMessage(srv);
            PropertySet psMime = null;
            ExtendedPropertyDefinition pr_flags = null;
            try
            {
                //Copia o conteudo usando o MimeContent
                psMime = new PropertySet(ItemSchema.MimeContent);
                msg.Load(psMime);

                emailMsg.MimeContent = msg.MimeContent;

                //Seta essa flag para que ele não salve a mensagem como se fosse Draft
                pr_flags = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
                emailMsg.SetExtendedProperty(pr_flags, "1");

                //Salva ID Antigo
                emailMsg.SetExtendedProperty(this.GetItemIDExtendedProperty(), msg.Id.UniqueId);

                emailMsg.Save(folderId);
            }
            //Quando a mensagem possui anexo, ele consegue copiar a mensagem so que joga um erro de Index OutOfRange
            catch (ArgumentOutOfRangeException)
            {

            }
            catch (Exception erro)
            {

                //Como o EWS gera erro ao tentar migrar mensagens de NDR ou confirmação de leitura, eu so gero o log de erro, mas 
                //Eu migro a mensagem mudando a flag de ItemClass
                if (string.Equals(erro.Message, "Operation would change object type, which is not permitted."))
                {
                    try
                    {
                        //Mudo a flag 
                        emailMsg.ItemClass = "IPM.Foo";

                        emailMsg.Save(folderId);
                    }
                    //Quando a mensagem possui anexo, ele consegue copiar a mensagem so que joga um erro de Index OutOfRange
                    catch (ArgumentOutOfRangeException)
                    {

                    }
                    catch (Exception erro1)
                    {
                        msg.Load();
                        this.WriteLog("Item Subject:" + msg.Subject + ": Idem Id: " + msg.Id.ToString() + "-Error:" + erro1.Message, targetMailbox);
                        countErrors++;
                    }
                }
                else
                {
                    msg.Load();
                    this.WriteLog("Item Subject:" + msg.Subject + ": Idem Id: " + msg.Id.ToString() + "-Error:" + erro.Message, targetMailbox);
                    countErrors++;
                }
            }
            finally
            {
                try
                {
                    if (emailMsg.MimeContent != null)
                        emailMsg.MimeContent.Content = null;
                    emailMsg = null;
                    msg = null;
                    GC.Collect();
                }
                catch { }
            }
        }



        private void SaveCalendar(ExchangeService srv, Appointment appt, FolderId folderId)
        {
            Appointment appointment = new Appointment(srv);

            try
            {
                //Copia todo o conteudo usando MimeContent
                PropertySet psMime = new PropertySet(ItemSchema.MimeContent);
                appt.Load(psMime);

                appointment.MimeContent = appt.MimeContent;

                appt.Load();

                appointment.Body = appt.Body;
                appointment.Categories = appt.Categories;
                appointment.Culture = appt.Culture;
                //appointment.DisplayCc = appt.DisplayCc;
                appointment.End = appt.End;
                try
                {
                    appointment.IsResponseRequested = appt.IsResponseRequested;
                }
                catch { }
                //appointment.MyResponseType = appt.MyResponseType;                


                //Esse objeto não existe no Exchange 2007, em uma migração ele falha
                try
                {
                    appointment.EndTimeZone = appt.EndTimeZone;
                }
                catch { }
                appointment.Importance = appt.Importance;
                appointment.InReplyTo = appt.InReplyTo;
                appointment.ItemClass = appt.ItemClass;
                appointment.LegacyFreeBusyStatus = appt.LegacyFreeBusyStatus;
                appointment.Location = appt.Location;
                appointment.MeetingWorkspaceUrl = appt.MeetingWorkspaceUrl;
                //appointment.MimeContent = appt.MimeContent;
                appointment.NetShowUrl = appt.NetShowUrl;
                //appointment.InternetMessageHeaders = appt.InternetMessageHeaders;

                try
                {
                    appointment.Organizer.Address = appt.Organizer.Address;
                }
                catch { }
                //appointment.OriginalStart.Add(appt.OriginalStart.ToUniversalTime);           
                //appointment.Preview = appt.Preview;
                appointment.Recurrence = appt.Recurrence;
                try
                {
                    appointment.ReminderDueBy = appt.ReminderDueBy;
                }
                catch { }
                appointment.ReminderMinutesBeforeStart = appt.ReminderMinutesBeforeStart;


                //ReadOnly Objects
                //appointment.Resources = appt.Resources;
                //appointment.RetentionDate = appt.RetentionDate;
                //appointment.Schema = appt.Schema;
                appointment.Sensitivity = appt.Sensitivity;
                try
                {
                    appointment.StartTimeZone = appt.StartTimeZone;
                }
                catch { }
                appointment.Subject = appt.Subject;
                //appointment.TimeZone = appt.TimeZone;
                //appointment.When = appt.When;
                appointment.IsAllDayEvent = appt.IsAllDayEvent;
                //appointment.IsAssociated = appt.IsAssociated;
                //appointment.IsDirty = appt.IsDirty;
                //appointment.IsDraft = appt.IsDraft;
                //appointment.IsFromMe = appt.IsFromMe;
                //appointment.IsMeeting = appt.IsMeeting;
                //appointment.IsNew = appt.IsNew;
                //appointment.IsRecurring. = appt.IsRecurring;
                appointment.IsReminderSet = appt.IsReminderSet;
                //appointment.IsResend = appt.IsResend;
                //appointment.IsUnmodified = appt.IsUnmodified;
                //appointment.JoinOnlineMeetingUrl = appointment.JoinOnlineMeetingUrl;

                appointment.Start = appt.Start;


                if (appt.Attachments.Count != 0)
                {

                    foreach (var att in appt.Attachments)
                    {
                        if (att is FileAttachment)
                        {
                            FileAttachment fileAtt = att as FileAttachment;
                            fileAtt.Load();
                            if (att.Name == null)
                            {
                                appointment.Attachments.AddFileAttachment("attachment", fileAtt.Content);
                            }
                            else
                            {
                                appointment.Attachments.AddFileAttachment(att.Name, fileAtt.Content);
                            }
                        }
                        if (att is ItemAttachment)
                        {
                            ItemAttachment itemAtt = att as ItemAttachment;
                            itemAtt.Load(new PropertySet(ItemSchema.MimeContent));

                            ItemAttachment<Item> item = appointment.Attachments.AddItemAttachment<Item>();
                            item.Name = itemAtt.Name;
                            item.Item.Subject = itemAtt.Item.Subject;
                            item.Item.MimeContent = itemAtt.Item.MimeContent;
                        }
                    }
                }

                if (appt.RequiredAttendees.Count != 0)
                {
                    foreach (Attendee obj in appt.RequiredAttendees)
                    {
                        appointment.RequiredAttendees.Add(obj.Address);
                    }
                }


                if (appt.OptionalAttendees.Count != 0)
                {
                    foreach (Attendee obj in appt.OptionalAttendees)
                    {
                        appointment.OptionalAttendees.Add(obj.Address);
                    }
                }

                //Salva ID Antigo
                appointment.SetExtendedProperty(this.GetItemIDExtendedProperty(), appt.Id.UniqueId);


                //Setando o Ateendees//Esse guid é usado para todos os RequiredAttendee
                //Tive que fazer isso porque ao criar um novo item no calendario ele enviava novamente os invites
                //Outro problema que tive foi tratar invites enviados por outras pessoas, por isso tive que manter
                //algumas flags das mensagens. Para visualizar todas as propriedades estendidas de uma mensagem
                //eu instalei o Outlook Spy que permite visualizar esses itens
                Guid guid = new Guid("{00062002-0000-0000-C000-000000000046}");
                ExtendedPropertyDefinition toExtendedPropertyDefinition = new ExtendedPropertyDefinition(guid, 0x823B, MapiPropertyType.String);
                PropertySet psTo = new PropertySet(toExtendedPropertyDefinition);
                appt.Load(psTo);
                foreach (ExtendedProperty prop in appt.ExtendedProperties)
                {
                    if (prop.Value != null)
                        appointment.SetExtendedProperty(toExtendedPropertyDefinition, prop.Value.ToString());
                }
                //Setando o CCAttendeeString
                ExtendedPropertyDefinition ccExtendedPropertyDefinition = new ExtendedPropertyDefinition(guid, 0x823C, MapiPropertyType.String);
                PropertySet psCC = new PropertySet(ccExtendedPropertyDefinition);
                appt.Load(psCC);

                foreach (ExtendedProperty prop in appt.ExtendedProperties)
                {
                    if (prop.Value != null)
                        appointment.SetExtendedProperty(ccExtendedPropertyDefinition, prop.Value.ToString());
                }

                //Adicionando Convidados

                ExtendedPropertyDefinition extendedPropertyDefinition = new ExtendedPropertyDefinition(guid, 0x8238, MapiPropertyType.String);
                PropertySet ps = new PropertySet(extendedPropertyDefinition);
                appt.Load(ps);
                //Le propriedade extentidade do RequiredAttendees
                foreach (ExtendedProperty prop in appt.ExtendedProperties)
                {
                    if (prop.Value != null)
                        appointment.SetExtendedProperty(extendedPropertyDefinition, prop.Value.ToString());
                }
                //Adicionando ResponseStatus
                ExtendedPropertyDefinition respStatusExtendedPropertyDefinition = new ExtendedPropertyDefinition(guid, 0x8218, MapiPropertyType.Integer);
                PropertySet psRespStatus = new PropertySet(respStatusExtendedPropertyDefinition);
                appt.Load(psRespStatus);
                foreach (ExtendedProperty prop in appt.ExtendedProperties)
                {
                    if (prop.Value != null)
                        appointment.SetExtendedProperty(respStatusExtendedPropertyDefinition, prop.Value.ToString());
                }
                //Adicionando Appointment State Flags
                ExtendedPropertyDefinition appStatedFlagExtendedPropertyDefinition = new ExtendedPropertyDefinition(guid, 0x8217, MapiPropertyType.Integer);
                PropertySet psAppStateFlag = new PropertySet(appStatedFlagExtendedPropertyDefinition);
                appt.Load(psAppStateFlag);
                foreach (ExtendedProperty prop in appt.ExtendedProperties)
                {
                    if (prop.Value != null)
                        appointment.SetExtendedProperty(appStatedFlagExtendedPropertyDefinition, prop.Value.ToString());
                }

                //Adicionando ResponseState
                Guid respStateGuid = new Guid("{6ED8DA90-450B-101B-98DA-00AA003F1305}");
                ExtendedPropertyDefinition respStateExtendedPropertyDefinition = new ExtendedPropertyDefinition(respStateGuid, 0x21, MapiPropertyType.Integer);
                PropertySet psRespState = new PropertySet(respStateExtendedPropertyDefinition);
                appt.Load(psRespState);
                foreach (ExtendedProperty prop in appt.ExtendedProperties)
                {
                    if (prop.Value != null)
                        appointment.SetExtendedProperty(respStateExtendedPropertyDefinition, prop.Value.ToString());
                }

                appointment.Save(folderId, SendInvitationsMode.SendToNone);
            }
            catch (Exception erro)
            {
                appt.Load();
                this.WriteLog("Item Subject:" + appt.Subject + ": Idem Id: " + appt.Id.ToString() + "-Error:" + erro.Message, targetMailbox);
                countErrors++;
            }
            finally
            {
                try
                {
                    //Libero recursos, tambem forço o garbage collector para liberar memoria
                    //Tive problemas de consumo de memoria, somente com o GC consegui resolver isso
                    appt = null;
                    appointment = null;
                    GC.Collect();
                }
                catch { }
            }
        }

        private void SaveTask(ExchangeService srv, Task task, FolderId folderId)
        {
            Task cpTask = new Task(srv);

            try
            {
                task.Load();

                cpTask.ActualWork = task.ActualWork;

                //cpTask.AssignedTime = task.AssignedTime;          

                cpTask.BillingInformation = task.BillingInformation;
                cpTask.Body = task.Body;
                cpTask.Categories = task.Categories;
                //cpTask.ChangeCount = task.ChangeCount;            
                cpTask.Companies = task.Companies;
                cpTask.CompleteDate = task.CompleteDate;
                cpTask.Contacts = task.Contacts;
                cpTask.Culture = task.Culture;
                //cpTask.DateTimeCreated = task.DateTimeCreated;
                //cpTask.DateTimeReceived = task.DateTimeReceived;
                //cpTask.DateTimeSent = task.DateTimeSent;
                //cpTask.DisplayCc = task.DisplayCc;
                //cpTask.DisplayTo = task.DisplayTo;
                cpTask.InReplyTo = task.InReplyTo;
                cpTask.Importance = task.Importance;
                cpTask.Mileage = task.Mileage;
                //cpTask.MimeContent = task.MimeContent;
                //cpTask.Mode = task.Mode;
                //cpTask.NormalizedBody = task.NormalizedBody;                                   
                //cpTask.Owner = task.Owner;
                cpTask.PercentComplete = task.PercentComplete;
                cpTask.Recurrence = task.Recurrence;
                cpTask.ReminderDueBy = task.ReminderDueBy;
                cpTask.ReminderMinutesBeforeStart = task.ReminderMinutesBeforeStart;
                cpTask.Sensitivity = task.Sensitivity;
                cpTask.StartDate = task.StartDate;
                cpTask.Status = task.Status;
                //cpTask.StatusDescription = task.StatusDescription;
                cpTask.Subject = task.Subject;
                cpTask.TotalWork = task.TotalWork;
                //cpTask.UniqueBody.Text= task.UniqueBody.BodyType;                        


                //Tratando Anexos

                if (task.Attachments.Count != 0)
                {
                    foreach (var att in task.Attachments)
                    {
                        if (att is FileAttachment)
                        {
                            FileAttachment fileAtt = att as FileAttachment;
                            fileAtt.Load();
                            if (att.Name == null)
                            {
                                cpTask.Attachments.AddFileAttachment("attachment", fileAtt.Content);
                            }
                            else
                            {
                                cpTask.Attachments.AddFileAttachment(att.Name, fileAtt.Content);
                            }
                        }
                        if (att is ItemAttachment)
                        {
                            ItemAttachment itemAtt = att as ItemAttachment;
                            itemAtt.Load(new PropertySet(ItemSchema.MimeContent));

                            ItemAttachment<Item> item = cpTask.Attachments.AddItemAttachment<Item>();
                            item.Name = itemAtt.Name;
                            item.Item.Subject = itemAtt.Item.Subject;
                            item.Item.MimeContent = itemAtt.Item.MimeContent;
                        }
                    }
                }


                //Salva ID Antigo
                cpTask.SetExtendedProperty(this.GetItemIDExtendedProperty(), task.Id.UniqueId);


                cpTask.Save(folderId);
            }

            catch (Exception erro)
            {
                task.Load();
                this.WriteLog("Item Subject:" + task.Subject + ": Idem Id: " + task.Id.ToString() + "-Error:" + erro.Message, targetMailbox);
                countErrors++;
            }
            finally
            {
                try
                {
                    //Libero recursos, tambem forço o garbage collector para liberar memoria
                    //Tive problemas de consumo de memoria, somente com o GC consegui resolver isso
                    cpTask = null;
                    task = null;
                    GC.Collect();
                }
                catch { }
            }
        }

        private void SaveContact(ExchangeService srv, Contact contact, FolderId folderId)
        {
            Contact cpContact = new Contact(srv);
            try
            {
                contact.Load();

                cpContact.AssistantName = contact.AssistantName;
                cpContact.BusinessHomePage = contact.BusinessHomePage;

                //Copiando Notes(Detalhes)
                //A informacao esta no atributo Body
                cpContact.Body = contact.Body;
                try
                {
                    cpContact.Birthday = contact.Birthday;
                }
                catch { }
                //cpContact.ContactSource. = contact.ContactSource.Value;
                cpContact.Generation = contact.Generation;
                try
                {
                    cpContact.PostalAddressIndex = contact.PostalAddressIndex;
                }
                catch { }


                cpContact.Categories = contact.Categories;
                cpContact.Children = contact.Children;
                cpContact.Companies = contact.Companies;
                cpContact.CompanyName = contact.CompanyName;
                cpContact.Culture = contact.Culture;
                cpContact.Department = contact.Department;
                cpContact.DisplayName = contact.DisplayName;

                //Se existir algum email cadastrado ele copia
                try
                {
                    string email1 = contact.EmailAddresses[EmailAddressKey.EmailAddress1].ToString();
                    if (!string.IsNullOrEmpty(email1))
                        cpContact.EmailAddresses[EmailAddressKey.EmailAddress1] = new EmailAddress(email1);
                }
                catch { }
                //No exchange 2007 esses parametros não existem
                try
                {
                    string email2 = contact.EmailAddresses[EmailAddressKey.EmailAddress2].ToString();
                    if (!string.IsNullOrEmpty(email2))
                        cpContact.EmailAddresses[EmailAddressKey.EmailAddress2] = new EmailAddress(email2);
                }
                catch { }

                try
                {
                    string email3 = contact.EmailAddresses[EmailAddressKey.EmailAddress3].ToString();
                    if (!string.IsNullOrEmpty(email3))
                        cpContact.EmailAddresses[EmailAddressKey.EmailAddress3] = new EmailAddress(email3);
                }
                catch { }

                cpContact.FileAs = contact.FileAs;
                try
                {
                    cpContact.FileAsMapping = contact.FileAsMapping;
                }
                catch { }
                cpContact.Generation = contact.Generation;
                cpContact.GivenName = contact.GivenName;


                try
                {
                    cpContact.ImAddresses[ImAddressKey.ImAddress1] = contact.ImAddresses[ImAddressKey.ImAddress1];
                }
                catch { }
                //Comentado porque gera erro no catalogo
                //cpContact.ImAddresses[ImAddressKey.ImAddress2] = contact.ImAddresses[ImAddressKey.ImAddress2];
                //cpContact.ImAddresses[ImAddressKey.ImAddress3] = contact.ImAddresses[ImAddressKey.ImAddress3];

                cpContact.Importance = contact.Importance;
                cpContact.Initials = contact.Initials;
                cpContact.InReplyTo = contact.InReplyTo;
                cpContact.JobTitle = contact.JobTitle;
                cpContact.Manager = contact.Manager;
                cpContact.MiddleName = contact.MiddleName;
                cpContact.Mileage = contact.Mileage;
                cpContact.NickName = contact.NickName;
                cpContact.OfficeLocation = contact.OfficeLocation;

                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.AssistantPhone] = contact.PhoneNumbers[PhoneNumberKey.AssistantPhone];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.BusinessPhone] = contact.PhoneNumbers[PhoneNumberKey.BusinessPhone];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.BusinessPhone2] = contact.PhoneNumbers[PhoneNumberKey.BusinessPhone2];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.Callback] = contact.PhoneNumbers[PhoneNumberKey.Callback];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.CarPhone] = contact.PhoneNumbers[PhoneNumberKey.CarPhone];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.CompanyMainPhone] = contact.PhoneNumbers[PhoneNumberKey.CompanyMainPhone];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.HomePhone] = contact.PhoneNumbers[PhoneNumberKey.HomePhone];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.HomePhone2] = contact.PhoneNumbers[PhoneNumberKey.HomePhone2];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.Isdn] = contact.PhoneNumbers[PhoneNumberKey.Isdn];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.MobilePhone] = contact.PhoneNumbers[PhoneNumberKey.MobilePhone];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.OtherTelephone] = contact.PhoneNumbers[PhoneNumberKey.OtherTelephone];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.Pager] = contact.PhoneNumbers[PhoneNumberKey.Pager];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.PrimaryPhone] = contact.PhoneNumbers[PhoneNumberKey.PrimaryPhone];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.RadioPhone] = contact.PhoneNumbers[PhoneNumberKey.RadioPhone];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.Telex] = contact.PhoneNumbers[PhoneNumberKey.Telex];
                }
                catch { }
                try
                {
                    cpContact.PhoneNumbers[PhoneNumberKey.TtyTddPhone] = contact.PhoneNumbers[PhoneNumberKey.TtyTddPhone];
                }
                catch { }
                try
                {
                    cpContact.PhysicalAddresses[PhysicalAddressKey.Business] = contact.PhysicalAddresses[PhysicalAddressKey.Business];
                }
                catch { }
                try
                {
                    cpContact.PhysicalAddresses[PhysicalAddressKey.Home] = contact.PhysicalAddresses[PhysicalAddressKey.Home];
                }
                catch { }
                try
                {
                    cpContact.PhysicalAddresses[PhysicalAddressKey.Other] = contact.PhysicalAddresses[PhysicalAddressKey.Other];
                }
                catch { }


                cpContact.Profession = contact.Profession;
                cpContact.SpouseName = contact.SpouseName;
                cpContact.Subject = contact.Subject;
                cpContact.Surname = contact.Surname;


                //Tratando Anexos    

                if (contact.Attachments.Count != 0)
                {
                    foreach (var att in contact.Attachments)
                    {
                        if (att is FileAttachment)
                        {
                            FileAttachment fileAtt = att as FileAttachment;
                            fileAtt.Load();
                            if (att.Name == null)
                            {
                                cpContact.Attachments.AddFileAttachment("attachment", fileAtt.Content);
                            }
                            else
                            {
                                cpContact.Attachments.AddFileAttachment(att.Name, fileAtt.Content);
                            }
                        }
                        if (att is ItemAttachment)
                        {
                            ItemAttachment itemAtt = att as ItemAttachment;
                            itemAtt.Load(new PropertySet(ItemSchema.MimeContent));

                            ItemAttachment<Item> item = cpContact.Attachments.AddItemAttachment<Item>();
                            item.Name = itemAtt.Name;
                            item.Item.Subject = itemAtt.Item.Subject;
                            item.Item.MimeContent = itemAtt.Item.MimeContent;
                        }
                    }
                }

                //Salva OlderID
                cpContact.SetExtendedProperty(this.GetItemIDExtendedProperty(), contact.Id.UniqueId);


                cpContact.Save(folderId);
            }

            catch (Exception erro)
            {
                this.WriteLog("Item Subject:" + contact.Subject + ": Idem Id: " + contact.Id.ToString() + "-Error:" + erro.Message, targetMailbox);
                countErrors++;
            }
            finally
            {
                try
                {
                    //Libero recursos, tambem forço o garbage collector para liberar memoria
                    //Tive problemas de consumo de memoria, somente com o GC consegui resolver isso
                    cpContact = null;
                    contact = null;
                    GC.Collect();
                }
                catch { }
            }
        }

        private List<Item> GetFolderItems(ExchangeService svc, FolderId folderId)
        {
            ItemView view = new ItemView(int.MaxValue);
            //Adiciona na busca a ExtendedProperty onde fica armazenado o antigo ItemID para não duplicar objetos
            view.PropertySet = new PropertySet(this.GetItemIDExtendedProperty());
            //WorkAround para bloqueio das Throlling Policies
            FindItemsResults<Item> itemResults = null;
            List<Item> list = new List<Item>();

            do
            {
                itemResults = svc.FindItems(folderId, view);
                foreach (Item item in itemResults)
                {
                    list.Add(item);
                }
                view.Offset += itemResults.Items.Count;
            } while (itemResults.MoreAvailable);

            return list;
        }


        private List<ItemIdentity> GetIdsMigratedItems(ExchangeService svc, WellKnownFolderName rootFolder)
        {
            //Metodo retorna o ID antigo da mensagem para fazer de/para na migracao
            List<ItemIdentity> itemResults = new List<ItemIdentity>();
            List<FolderIdentity> pathList = new List<FolderIdentity>();
            FolderView folderView = new FolderView(int.MaxValue);
            folderView.Traversal = FolderTraversal.Deep;

            foreach (Folder childFolder in svc.FindFolders(WellKnownFolderName.MsgFolderRoot, folderView))
            {
                ItemView itemView = new ItemView(int.MaxValue);
                //Adiciona na busca a ExtendedProperty onde fica armazenado o antigo ItemID para não duplicar objetos
                itemView.PropertySet = new PropertySet(this.GetItemIDExtendedProperty());
                //Como as Throlling Policies do Exchange limitam a Find de 1000 itens
                //Precisamos fazer um workaround no código ele continua a busca até terminar todos os itens
                //Por isso o uso do Do {}while()
                FindItemsResults<Item> folderitens = null;

                do
                {
                    folderitens = svc.FindItems(childFolder.Id, itemView);

                    foreach (Item i in folderitens)
                    {
                        //Le o Extented Attribute que possui o ID antigo da mensagem
                        foreach (ExtendedProperty exProp in i.ExtendedProperties)
                        {
                            string propName = exProp.PropertyDefinition.Name;

                            if (propName.ToLower() == "olderid")
                            {
                                //Adiciona o ID atual e o ID antigo que foi migrado
                                string newID = i.Id.ToString();
                                string oldID = exProp.Value.ToString();
                                if (string.IsNullOrEmpty(oldID))
                                {
                                    throw new Exception("Older ID of a already migrated item cannot be null, it can cause duplicated itens!!");
                                }

                                itemResults.Add(new ItemIdentity(newID.Trim(), oldID.Trim()));
                            }
                        }
                    }
                    itemView.Offset += folderitens.Items.Count;
                } while (folderitens.MoreAvailable);
            }

            return itemResults;
        }


        private ExtendedPropertyDefinition GetItemIDExtendedProperty()
        {
            Guid MyPropertySetId = new Guid("{C11FF724-AA03-4555-9952-8FA248A11C3E}");

            ExtendedPropertyDefinition extendedPropertyDefinition = new ExtendedPropertyDefinition(MyPropertySetId, "OlderID", MapiPropertyType.String);

            return extendedPropertyDefinition;
        }

        private static bool OnValidationCallback(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }
    }

    public class FolderIdentity
    {
        public string FolderPath { get; set; }
        public string FolderType { get; set; }
        public FolderId FolderID { get; set; }
        public int CountItems { get; set; }

        public FolderIdentity(string folderPath, string folderType, int countItems)
        {
            this.FolderPath = folderPath;
            this.FolderType = folderType;
            this.CountItems = countItems;
        }

        public FolderIdentity(string folderPath, string folderType, FolderId folderId, int countItems)
        {
            this.FolderPath = folderPath;
            this.FolderType = folderType;
            this.FolderID = folderId;
            this.CountItems = countItems;
        }
    }

    public class ItemIdentity
    {
        public string AppNewId { get; set; }
        public string AppOldId { get; set; }

        public ItemIdentity(string appNewId, string appOldId)
        {
            this.AppNewId = appNewId;
            this.AppOldId = appOldId;
        }
    }
}
