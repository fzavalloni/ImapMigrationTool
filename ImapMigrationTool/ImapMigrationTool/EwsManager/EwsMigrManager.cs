using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Net.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using ImapMigrationTool.Entities;
using System.Windows.Forms;

namespace ImapMigrationTool.EwsManager
{
    public class EwsMigrManager
    {

        bool ewsMessageCopy = false;
        bool ewsContactsCopy = false;
        bool ewsCalendarCopy = false;
        bool ewsTasksCopy = false;
        bool logEnabled = false;
        bool ewsInboxRuleCopy = false;
        bool removeDuplicate = false;
        string targetMailbox = string.Empty;
        int countErrors = 0;

        private EwsMgrManagerBase ewsManagerBase;
        private EwsHelper helper;

        public EwsMigrManager()
        {
            ewsManagerBase = new EwsMgrManagerBase();
            helper = new EwsHelper();
        }

        #region Public Methods

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
                    exSrcServer = ewsManagerBase.GetService(data.SourceAdminUser, data.SourceAdminPassword, srcUserMail, ewsSrcUrl, exSourceVersion);
                }
                else
                {
                    exSrcServer = ewsManagerBase.GetService(srcUserMail, password, ewsSrcUrl, exSourceVersion);
                }

                string result = GetMailboxSize(exSrcServer);

                //Recuperando idioma da pasta

                result = string.Format("{0} - {1}", result, GetMailboxLanguage(exSrcServer, srcUserMail));

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
                //Recupera informacoes de migracao de objetos
                ewsMessageCopy = data.EWSMessage;
                ewsContactsCopy = data.EWSContacts;
                ewsCalendarCopy = data.EWSCalendar;
                ewsTasksCopy = data.EWSTasks;
                ewsInboxRuleCopy = data.EWSInboxRules;
                removeDuplicate = data.RemoveDuplicate;

                targetMailbox = tgtUserMail;

                this.WriteLog("Process started: " + DateTime.Now, targetMailbox);

                //Recupera versão do Exchange e converte para Enum
                ExchangeVersion exSourceVersion = (ExchangeVersion)Enum.Parse(typeof(ExchangeVersion), data.SrcServerVersion);
                ExchangeVersion exTargetVersion = (ExchangeVersion)Enum.Parse(typeof(ExchangeVersion), data.TgtServerVersion);

                //Montando serviços
                string ewsSrcUrl = string.Format("https://{0}/ews/exchange.asmx", data.SourceServer);
                string ewsTgtUrl = string.Format("https://{0}/ews/exchange.asmx", data.TargetServer);
                ExchangeService exTgtServer = ewsManagerBase.GetService(data.TargetAdminUser, data.TargetAdminPassword, tgtUserMail, ewsTgtUrl, exTargetVersion);
                ExchangeService exSrcServer = null;

                if (isAdminMigration)
                {
                    exSrcServer = ewsManagerBase.GetService(data.SourceAdminUser, data.SourceAdminPassword, srcUserMail, ewsSrcUrl, exSourceVersion);
                }
                else
                {
                    exSrcServer = ewsManagerBase.GetService(srcUserMail, password, ewsSrcUrl, exSourceVersion);
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
                    if (ewsInboxRuleCopy)
                    {
                        //Migrando Inbox Rules
                        Tools.SetRowValue(ref row, EColumns.results, "Migrating Inbox Rules....");
                        this.MigrateInboxRules(exSrcServer, exTgtServer);
                        //Migrando Mensagens                
                        Tools.SetRowValue(ref row, EColumns.results, string.Empty);
                    }
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

        #endregion

        private void MigrateItens(ref DataGridViewRow row, ExchangeService srcService, ExchangeService tgtService)
        {
            //Carrega lista de pastas que serao migradas fazendo de/para dos Ids das pastas.
            List<FolderIdentity> folderSourceList = GetFoldersIds(srcService);
            List<FolderIdentity> folderTargetList = GetFoldersIds(tgtService);
            int totalMessagesToBeMigrated = helper.SumFolderCountMessages(folderSourceList);

            Tools.SetRowValue(ref row, EColumns.steps, "0 %");

            int processedItemCount = 0;
            int totalCopiedFolderitemCount = 0;
            int countMeetingRequestObjects = 0;
            int folderProcessedErrors = 0;

            foreach (FolderIdentity folderSource in folderSourceList)
            {
                bool skipFailedFolder = false;
                int copiedFolderitemCount = 0;

                //Recupera objetos da origem

                List<Item> contactsListSrc = this.GetFolderItems(srcService, folderSource.FolderID);


                string folderProgress = "Processing folder: " + folderSource.FolderPath;
                Tools.SetRowValue(ref row, EColumns.results, folderProgress);
                this.WriteLog("Processing folder: " + folderSource.FolderPath, targetMailbox);

                //Recupera objetos do destino
                //Recupera ID da mesma pasta de destino

                FolderId tgtFolderId = helper.PickCurrentObjInList(folderSource, folderTargetList);

                //Carrega todos os Ids dos Objetos que já foram migrados
                //Em caso de falha, ele pula o diretório
                List<ItemIdentity> migratedListIds = null;
                try
                {
                    migratedListIds = this.GetFolderIdsMigratedItems(tgtService, tgtFolderId);
                }
                catch (Exception err)
                {
                    this.WriteLog("Failed to process folder: " + err.Message, targetMailbox);
                    countErrors++;
                    folderProcessedErrors++;
                    skipFailedFolder = true;
                }
                //Se nao houve falha na busca dos IDs, processa a copia das mensagens
                if (!skipFailedFolder)
                {
                    foreach (Item obj in contactsListSrc)
                    {
                        processedItemCount++;

                        //int percentCompleted = ((count / totalMessagesToBeMigrated) * 100);
                        int percentCompleted = (int)Math.Round((double)(100 * processedItemCount) / totalMessagesToBeMigrated);

                        Tools.SetRowValue(ref row, EColumns.steps, string.Format("{0} % - Errors {1}", percentCompleted, countErrors));
                        Tools.SetRowValue(ref row, EColumns.results, string.Format("{0} - Item {1}", folderProgress, processedItemCount));
                        //Caso o ID não exista no destino a cópia do objeto é feita
                        string id = obj.Id.UniqueId;

                        //Trata Appointments que tiveram alguma alteracao

                        if (ewsCalendarCopy)
                        {
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
                                            this.CopyCalendar(tgtService, oldApp, tgtFolderId);
                                            copiedFolderitemCount++;
                                        }
                                    }
                                }
                            }
                        }


                        if (!migratedListIds.Exists(i => i.AppOldId == id))
                        {
                            try
                            {

                                if (obj is Contact)
                                {
                                    if (ewsContactsCopy)
                                    {
                                        Contact contact = obj as Contact;
                                        this.CopyContact(tgtService, contact, tgtFolderId);
                                        copiedFolderitemCount++;
                                    }
                                }

                                if (obj is ContactGroup)
                                {
                                    if (ewsContactsCopy)
                                    {
                                        ContactGroup contactGroup = obj as ContactGroup;
                                        this.CopyContactGroup(tgtService, contactGroup, tgtFolderId);
                                        copiedFolderitemCount++;
                                    }
                                }

                                if (obj is Appointment)
                                {
                                    if (ewsCalendarCopy)
                                    {
                                        Appointment appt = obj as Appointment;
                                        this.CopyCalendar(tgtService, appt, tgtFolderId);
                                        copiedFolderitemCount++;
                                    }
                                }

                                if (obj is Task)
                                {
                                    if (ewsTasksCopy)
                                    {
                                        Task task = obj as Task;
                                        this.CopyTask(tgtService, task, tgtFolderId);
                                        copiedFolderitemCount++;
                                    }
                                }

                                if (obj is EmailMessage)
                                {

                                    if (ewsMessageCopy)
                                    {
                                        if ((obj is MeetingRequest) || (obj is MeetingCancellation) || (obj is MeetingResponse))
                                        {
                                            //Itens de Meeting Request não são migrados, por isso só aparecem no report.
                                            countMeetingRequestObjects++;
                                        }
                                        else
                                        {
                                            EmailMessage msg = obj as EmailMessage;
                                            this.CopyMessage(tgtService, msg, tgtFolderId);
                                            copiedFolderitemCount++;
                                        }
                                    }

                                }
                            }
                            catch (Exception err)
                            {
                                this.WriteLog(string.Format("Corrupted Item: {0} - Error: {1}", obj.Id, err.Message), targetMailbox);
                                countErrors++;
                            }
                        }
                    }


                    //Carrega todos os Ids dos objetos que da origem para remover mensagens apagadas na origem                    
                    if (removeDuplicate)
                    {
                        int deletedObjects = 0;
                        Tools.SetRowValue(ref row, EColumns.results, string.Format("{0} | Removing removed/moved itens in target folder", folderSource.FolderPath));

                        deletedObjects = this.RemoveMovedOrDeletedItens(tgtService, contactsListSrc, migratedListIds);
                        this.WriteLog(string.Format("{0} - Deleted Duplicated Objects: {1}", folderSource.FolderPath, deletedObjects), targetMailbox);
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
            this.WriteLog(string.Format("Failed folders: {0}", folderProcessedErrors), targetMailbox);
            this.WriteLog(string.Format("Ignored Deleted Itens Folder: {0}", helper.GetDeletedItensCount(srcService)), targetMailbox);
            this.WriteLog(string.Format("Ignored Meeting Request Itens: {0} itens", countMeetingRequestObjects), targetMailbox);
            Tools.SetRowValue(ref row, EColumns.steps, "Errors: " + countErrors);
            this.WriteLog("Errors: " + countErrors, targetMailbox);
        }

        private int RemoveMovedOrDeletedItens(ExchangeService tgtService, List<Item> sourceItens, List<ItemIdentity> targetItens)
        {
            int deletedObjects = 0;
            foreach (ItemIdentity tgtItem in targetItens)
            {
                if (!sourceItens.Exists(i => i.Id.UniqueId == tgtItem.AppOldId))
                {
                    Item removedItem = Item.Bind(tgtService, tgtItem.AppNewId);
                    removedItem.Delete(DeleteMode.HardDelete, true);
                    deletedObjects++;
                }
            }
            return deletedObjects;
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
                        bool isruleAlreadyCreated = helper.IsRuleAlreadyCreated(ruleSource.DisplayName, ruleCollectionTarget);
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

        private void SubscribeFolders(ref DataGridViewRow row, ExchangeService exSrcServer, ExchangeService exTgtServer)
        {
            //Recupera a lista de todas as pasta na origem e destino
            List<FolderIdentity> folderListSourceSrv = GetFolderHierarchy(exSrcServer);
            List<FolderIdentity> folderListTargetSrv = GetFolderHierarchy(exTgtServer);
            //Gera uma nova lista apenas com as pastas que precisam ser criadas
            List<FolderIdentity> diffList = helper.RemoveFoldersAlreadyExists(folderListSourceSrv, folderListTargetSrv);

            FolderId lastFolderId = null;
            EwsFolderMgr fMgr = new EwsFolderMgr();

            foreach (FolderIdentity path in diffList)
            {
                Tools.SetRowValue(ref row, EColumns.results, "Creating folder hierarchy: " + path.FolderPath);

                if (path.FolderPath.Contains("\\"))
                {
                    string[] folders = path.FolderPath.Split('\\');

                    for (int i = 0; i < folders.Length; i++)
                    {

                        if (i == 0)
                        {
                            //Cria a pasta na raiz
                            fMgr.CreateNewFolder(exTgtServer, folders[i], WellKnownFolderName.MsgFolderRoot, path.FolderType);
                            lastFolderId = GetFolderId(exTgtServer, WellKnownFolderName.MsgFolderRoot, folders[i]);
                        }
                        else
                        {
                            //Cria as pastas filhas 
                            fMgr.CreateNewFolder(exTgtServer, folders[i], lastFolderId, path.FolderType);
                            lastFolderId = GetFolderId(exTgtServer, lastFolderId, folders[i]);
                        }
                    }
                }
                else
                {
                    //Create Folder no Root da pasta
                    fMgr.CreateNewFolder(exTgtServer, path.FolderPath, WellKnownFolderName.MsgFolderRoot, path.FolderType);
                }
            }
        }

        #region Copy Item Methods
        private void CopyMessage(ExchangeService srv, EmailMessage msg, FolderId folderId)
        {
            EmailMessage emailMsg = null;
            PropertySet psIsRead = null;
            ExtendedPropertyDefinition pr_message_delivery_time = null;
            ExtendedPropertyDefinition pr_client_submit_time = null;
            ExtendedPropertyDefinition pr_flags = null;
            ExtendedPropertyDefinition PidTagFlagStatus = null;
            ExtendedPropertyDefinition PidLidTaskDueDate = null;
            ExtendedPropertyDefinition PidLidTaskStartDate = null;
            ExtendedPropertyDefinition PidLidCommonStart = null;
            ExtendedPropertyDefinition PidLidCommonEnd = null;
            ExtendedPropertyDefinition PidTagFollowupIcon = null;

            try
            {
                emailMsg = new EmailMessage(srv);
                //Copia o conteudo usando o MimeContent                
                msg.Load(new PropertySet(ItemSchema.MimeContent));

                emailMsg.MimeContent = msg.MimeContent;

                //WorkAround necessario para o Exchange 2013, onde ele nao
                //reconhecia mensagens marcadas como nao lidas
                psIsRead = new PropertySet(EmailMessageSchema.IsRead);
                msg.Load(psIsRead);
                if (!msg.IsRead)
                {
                    emailMsg.IsRead = false;
                }

                //Seta essa flag para que ele não salve a mensagem como se fosse Draft
                pr_flags = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
                emailMsg.SetExtendedProperty(pr_flags, "1");

                // Root flag status (none/set/completed)
                PidTagFlagStatus = new ExtendedPropertyDefinition(0x1090, MapiPropertyType.Integer);
                msg.Load(new PropertySet(PidTagFlagStatus));
                if (msg.ExtendedProperties.Count != 0)
                    emailMsg.SetExtendedProperty(PidTagFlagStatus, msg.ExtendedProperties[0].Value);

                // Start/end dates in local time
                PidLidTaskDueDate = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Task, 0x8105, MapiPropertyType.SystemTime);
                msg.Load(new PropertySet(PidLidTaskDueDate));
                if (msg.ExtendedProperties.Count != 0)
                    emailMsg.SetExtendedProperty(PidLidTaskDueDate, msg.ExtendedProperties[0].Value);

                PidLidTaskStartDate = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Task, 0x8104, MapiPropertyType.SystemTime);
                msg.Load(new PropertySet(PidLidTaskStartDate));
                if (msg.ExtendedProperties.Count != 0)
                    emailMsg.SetExtendedProperty(PidLidTaskStartDate, msg.ExtendedProperties[0].Value);

                // Start/end dates in UTC time
                PidLidCommonStart = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Common, 0x8516, MapiPropertyType.SystemTime);
                msg.Load(new PropertySet(PidLidCommonStart));
                if (msg.ExtendedProperties.Count != 0)
                    emailMsg.SetExtendedProperty(PidLidCommonStart, msg.ExtendedProperties[0].Value);

                PidLidCommonEnd = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Common, 0x8517, MapiPropertyType.SystemTime);
                msg.Load(new PropertySet(PidLidCommonEnd));
                if (msg.ExtendedProperties.Count != 0)
                    emailMsg.SetExtendedProperty(PidLidCommonEnd, msg.ExtendedProperties[0].Value);


                PidTagFollowupIcon = new ExtendedPropertyDefinition(0x1095, MapiPropertyType.Integer);
                msg.Load(new PropertySet(PidTagFollowupIcon));
                if (msg.ExtendedProperties.Count != 0)
                    emailMsg.SetExtendedProperty(PidTagFollowupIcon, msg.ExtendedProperties[0].Value);


                //Ajustando manualmente o atributo delivery time, isso porque tivemos alguns casos, onde o Delivery Time migrava errado
                //Desta forma forçamos o envio no horario correto
                msg.Load();
                pr_message_delivery_time = new ExtendedPropertyDefinition(0x0E06, MapiPropertyType.SystemTime);
                emailMsg.SetExtendedProperty(pr_message_delivery_time, msg.DateTimeReceived);
                //Ajustando atributo Submit Time
                pr_client_submit_time = new ExtendedPropertyDefinition(0x0039, MapiPropertyType.SystemTime);
                emailMsg.SetExtendedProperty(pr_client_submit_time, msg.DateTimeSent);

                //Se o tipo do Sender for SMTP, ele recupera o endereço de email
                //Isso foi uma correção para um problema de alguns Itens Enviados que 
                //vinham vazios no from

                if (string.Equals(msg.From.RoutingType, "SMTP", StringComparison.CurrentCultureIgnoreCase))
                {
                    emailMsg.From = new EmailAddress(msg.From.Name, msg.From.Address);
                }

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
                //Eu logo o erro mas não mostro 
                if (string.Equals(erro.Message, "Operation would change object type, which is not permitted."))
                {
                    msg.Load();
                    this.WriteLog("Item Subject:" + msg.Subject + ": Idem Id: " + msg.Id.ToString() + "-Error: Bounce Messages are not migrated.", targetMailbox);
                }
                else
                {
                    msg.Load();
                    this.WriteLog("Item Subject:" + msg.Subject + ": Idem Id: " + msg.Id.ToString() + ":Item Size: " + EwsHelper.CovertByteToMB(msg.Size) + " MB - Error:" + erro.Message, targetMailbox);
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
                    //GC.Collect();
                }
                catch { }
            }
        }

        private void CopyCalendar(ExchangeService srv, Appointment appt, FolderId folderId)
        {
            Appointment appointment = new Appointment(srv);

            try
            {
                //Copia todo o conteudo usando MimeContent
                PropertySet psMime = new PropertySet(ItemSchema.MimeContent);
                appt.Load(psMime);

                appointment.MimeContent = appt.MimeContent;

                appt.Load();

                try
                {
                    appointment.Body = appt.Body;
                }
                catch { }
                try
                {
                    appointment.Categories = appt.Categories;
                }
                catch { }
                try
                {
                    appointment.Culture = appt.Culture;
                }
                catch { }
                //appointment.DisplayCc = appt.DisplayCc;
                try
                {
                    appointment.End = appt.End;
                }
                catch { }
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
                try
                {
                    appointment.Importance = appt.Importance;
                }
                catch { }
                try
                {
                    appointment.InReplyTo = appt.InReplyTo;
                }
                catch { }
                try
                {
                    appointment.ItemClass = appt.ItemClass;
                }
                catch { }
                try
                {
                    appointment.LegacyFreeBusyStatus = appt.LegacyFreeBusyStatus;
                }
                catch { }
                try
                {
                    appointment.Location = appt.Location;
                }
                catch { }
                try
                {
                    appointment.MeetingWorkspaceUrl = appt.MeetingWorkspaceUrl;
                }
                catch { }
                //appointment.MimeContent = appt.MimeContent;
                try
                {
                    appointment.NetShowUrl = appt.NetShowUrl;
                }
                catch { }
                //appointment.InternetMessageHeaders = appt.InternetMessageHeaders;

                try
                {
                    appointment.Organizer.Address = appt.Organizer.Address;
                }
                catch { }
                //appointment.OriginalStart.Add(appt.OriginalStart.ToUniversalTime);           
                //appointment.Preview = appt.Preview;
                try
                {
                    appointment.Recurrence = appt.Recurrence;
                }
                catch { }
                try
                {
                    appointment.ReminderDueBy = appt.ReminderDueBy;
                }
                catch { }
                try
                {
                    appointment.ReminderMinutesBeforeStart = appt.ReminderMinutesBeforeStart;
                }
                catch { }


                //ReadOnly Objects
                //appointment.Resources = appt.Resources;
                //appointment.RetentionDate = appt.RetentionDate;
                //appointment.Schema = appt.Schema;
                try
                {
                    appointment.Sensitivity = appt.Sensitivity;
                }
                catch { }
                try
                {
                    appointment.StartTimeZone = appt.StartTimeZone;
                }
                catch { }
                try
                {
                    appointment.Subject = appt.Subject;
                }
                catch { }
                //appointment.TimeZone = appt.TimeZone;
                //appointment.When = appt.When;
                try
                {
                    appointment.IsAllDayEvent = appt.IsAllDayEvent;
                }
                catch { }
                //appointment.IsAssociated = appt.IsAssociated;
                //appointment.IsDirty = appt.IsDirty;
                //appointment.IsDraft = appt.IsDraft;
                //appointment.IsFromMe = appt.IsFromMe;
                //appointment.IsMeeting = appt.IsMeeting;
                //appointment.IsNew = appt.IsNew;
                //appointment.IsRecurring. = appt.IsRecurring;
                try
                {
                    appointment.IsReminderSet = appt.IsReminderSet;
                }
                catch { }
                //appointment.IsResend = appt.IsResend;
                //appointment.IsUnmodified = appt.IsUnmodified;
                //appointment.JoinOnlineMeetingUrl = appointment.JoinOnlineMeetingUrl;
                try
                {
                    appointment.Start = appt.Start;
                }
                catch { }


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
                        appointment.RequiredAttendees.Add(helper.GetCorrectedAttendee(obj));
                    }
                }


                if (appt.OptionalAttendees.Count != 0)
                {
                    foreach (Attendee obj in appt.OptionalAttendees)
                    {
                        appointment.OptionalAttendees.Add(helper.GetCorrectedAttendee(obj));
                    }
                }

                if (appt.Resources.Count != 0)
                {
                    foreach (Attendee obj in appt.Resources)
                    {
                        appointment.Resources.Add(helper.GetCorrectedAttendee(obj));
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
                //FOI REMOVIDO POIS ESTAVA DANDO ERROS COM O EXCHANGE 2013 SP1 CU10
                //ExtendedPropertyDefinition respStatusExtendedPropertyDefinition = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Appointment, 0x8218, MapiPropertyType.Integer);
                //PropertySet psRespStatus = new PropertySet(respStatusExtendedPropertyDefinition);
                //appt.Load(psRespStatus);
                //foreach (ExtendedProperty prop in appt.ExtendedProperties)
                //{
                //    if (prop.Value != null)
                //    {
                //        appointment.SetExtendedProperty(respStatusExtendedPropertyDefinition, prop.Value.ToString());
                //    }
                //}

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
                this.WriteLog("Appointment Item:" + appt.Subject + ": Item Size: " + EwsHelper.CovertByteToMB(appt.Size) + " Idem Id: " + appt.Id.ToString() + "-Error:" + erro.Message, targetMailbox);
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
                    //GC.Collect();
                }
                catch { }
            }
        }

        private void CopyContactGroup(ExchangeService srv, ContactGroup contactGroup, FolderId folderId)
        {
            ContactGroup cpContactGroup = new ContactGroup(srv);

            try
            {
                contactGroup.Load();

                // cpContactGroup.DisplayCc = contactGroup.DisplayCc;
                try
                {
                    cpContactGroup.DisplayName = contactGroup.DisplayName;
                }
                catch { }
                //cpContactGroup.DisplayTo = contactGroup.DisplayTo;
                //cpContactGroup.FileAs = contactGroup.FileAs;

                try
                {
                    cpContactGroup.Body = contactGroup.Body;
                }
                catch { }

                try
                {
                    cpContactGroup.Flag = contactGroup.Flag;
                }
                catch { }
                //cpContactGroup.IconIndex = contactGroup.IconIndex;
                try
                {
                    cpContactGroup.Importance = contactGroup.Importance;
                }
                catch { }
                try
                {
                    cpContactGroup.IsReminderSet = contactGroup.IsReminderSet;
                }
                catch { }
                try
                {
                    cpContactGroup.ItemClass = contactGroup.ItemClass;
                }
                catch { }
                //cpContactGroup.LastModifiedName = contactGroup.LastModifiedName;
                if (contactGroup.Members.Count != 0)
                {
                    foreach (GroupMember obj in contactGroup.Members)
                    {
                        //Quando o tipo do contato e MailboxType.Contact ele é um contato do mesmo dominio
                        //Porem que nao existe mais, se tentar adicionar novamente mostra o erro abaixo
                        //MailboxType does not correspond to AD user recipient type
                        try
                        {
                            cpContactGroup.Members.Add(new GroupMember(obj.AddressInformation.Address.ToString(), MailboxType.OneOff));
                        }
                        catch { }
                    }
                }

                try
                {
                    cpContactGroup.MimeContent = contactGroup.MimeContent;
                }
                catch { }
                try
                {
                    cpContactGroup.Subject = contactGroup.Subject;
                }
                catch { }
                try
                {
                    cpContactGroup.PolicyTag = contactGroup.PolicyTag;
                }
                catch { }

                //Salva OlderID
                cpContactGroup.SetExtendedProperty(this.GetItemIDExtendedProperty(), contactGroup.Id.UniqueId);

                cpContactGroup.Save(folderId);
            }
            catch (Exception erro)
            {
                contactGroup.Load();
                this.WriteLog("Contact Group Item:" + contactGroup.Subject + ": Item Size " + EwsHelper.CovertByteToMB(contactGroup.Size) + ": Idem Id: " + contactGroup.Id.ToString() + "-Error:" + erro.Message, targetMailbox);
                countErrors++;
            }
            finally
            {
                try
                {
                    //Libero recursos, tambem forço o garbage collector para liberar memoria
                    //Tive problemas de consumo de memoria, somente com o GC consegui resolver isso
                    contactGroup = null;
                    cpContactGroup = null;
                    // GC.Collect();
                }
                catch { }
            }

        }

        private void CopyTask(ExchangeService srv, Task task, FolderId folderId)
        {
            Task cpTask = new Task(srv);

            try
            {
                task.Load();

                try
                {
                    cpTask.ActualWork = task.ActualWork;
                }
                catch { }

                //cpTask.AssignedTime = task.AssignedTime;          

                try
                {
                    cpTask.BillingInformation = task.BillingInformation;
                }
                catch { }
                try
                {
                    cpTask.Body = task.Body;
                }
                catch { }
                try
                {
                    cpTask.Categories = task.Categories;
                }
                catch { }
                //cpTask.ChangeCount = task.ChangeCount;            
                try
                {
                    cpTask.Companies = task.Companies;
                }
                catch { }
                try
                {
                    cpTask.CompleteDate = task.CompleteDate;
                }
                catch { }
                try
                {
                    cpTask.Contacts = task.Contacts;
                }
                catch { }
                try
                {
                    cpTask.Culture = task.Culture;
                }
                catch { }
                //cpTask.DateTimeCreated = task.DateTimeCreated;
                //cpTask.DateTimeReceived = task.DateTimeReceived;
                //cpTask.DateTimeSent = task.DateTimeSent;
                //cpTask.DisplayCc = task.DisplayCc;
                //cpTask.DisplayTo = task.DisplayTo;
                try
                {
                    cpTask.InReplyTo = task.InReplyTo;
                }
                catch { }
                try
                {
                    cpTask.Importance = task.Importance;
                }
                catch { }
                try
                {
                    cpTask.Mileage = task.Mileage;
                }
                catch { }
                //cpTask.MimeContent = task.MimeContent;
                //cpTask.Mode = task.Mode;
                //cpTask.NormalizedBody = task.NormalizedBody;                                   
                //cpTask.Owner = task.Owner;
                try
                {
                    cpTask.PercentComplete = task.PercentComplete;
                }
                catch { }
                try
                {
                    cpTask.Recurrence = task.Recurrence;
                }
                catch { }
                try
                {
                    cpTask.ReminderDueBy = task.ReminderDueBy;
                }
                catch { }
                try
                {
                    cpTask.ReminderMinutesBeforeStart = task.ReminderMinutesBeforeStart;
                }
                catch { }
                try
                {
                    cpTask.Sensitivity = task.Sensitivity;
                }
                catch { }
                try
                {
                    cpTask.StartDate = task.StartDate;
                }
                catch { }
                try
                {
                    cpTask.Status = task.Status;
                }
                catch { }
                //cpTask.StatusDescription = task.StatusDescription;
                try
                {
                    cpTask.Subject = task.Subject;
                }
                catch { }
                try
                {
                    cpTask.TotalWork = task.TotalWork;
                }
                catch { }
                try
                {
                    cpTask.IsReminderSet = task.IsReminderSet;
                }
                catch { }
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
                this.WriteLog("Task Item:" + task.Subject + ": Item Size: " + EwsHelper.CovertByteToMB(task.Size) + ": Idem Id: " + task.Id.ToString() + "-Error:" + erro.Message, targetMailbox);
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
                    //GC.Collect();
                }
                catch { }
            }
        }

        private void CopyContact(ExchangeService srv, Contact contact, FolderId folderId)
        {
            Contact cpContact = new Contact(srv);
            try
            {
                contact.Load();
                try
                {
                    cpContact.AssistantName = contact.AssistantName;
                }
                catch { }
                try
                {
                    cpContact.BusinessHomePage = contact.BusinessHomePage;
                }
                catch { }

                //Copiando Notes(Detalhes)
                //A informacao esta no atributo Body
                try
                {
                    cpContact.Body = contact.Body;
                }
                catch { }
                try
                {
                    cpContact.Birthday = contact.Birthday;
                }
                catch { }
                //cpContact.ContactSource. = contact.ContactSource.Value;
                try
                {
                    cpContact.Generation = contact.Generation;
                }
                catch { }
                try
                {
                    cpContact.PostalAddressIndex = contact.PostalAddressIndex;
                }
                catch { }

                try
                {
                    cpContact.Categories = contact.Categories;
                }
                catch { }
                try
                {
                    cpContact.Children = contact.Children;
                }
                catch { }
                try
                {
                    cpContact.Companies = contact.Companies;
                }
                catch { }
                try
                {
                    cpContact.CompanyName = contact.CompanyName;
                }
                catch { }
                try
                {
                    cpContact.Culture = contact.Culture;
                }
                catch { }
                try
                {
                    cpContact.Department = contact.Department;
                }
                catch { }
                try
                {
                    cpContact.DisplayName = contact.DisplayName;
                }
                catch { }

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
                try
                {
                    cpContact.FileAs = contact.FileAs;
                }
                catch { }
                try
                {
                    cpContact.FileAsMapping = contact.FileAsMapping;
                }
                catch { }
                try
                {
                    cpContact.Generation = contact.Generation;
                }
                catch { }
                try
                {
                    cpContact.GivenName = contact.GivenName;
                }
                catch { }


                try
                {
                    cpContact.ImAddresses[ImAddressKey.ImAddress1] = contact.ImAddresses[ImAddressKey.ImAddress1];
                }
                catch { }
                //Comentado porque gera erro no catalogo
                //cpContact.ImAddresses[ImAddressKey.ImAddress2] = contact.ImAddresses[ImAddressKey.ImAddress2];
                //cpContact.ImAddresses[ImAddressKey.ImAddress3] = contact.ImAddresses[ImAddressKey.ImAddress3];
                try
                {
                    cpContact.Importance = contact.Importance;
                }
                catch { }
                try
                {
                    cpContact.Initials = contact.Initials;
                }
                catch { }
                try
                {
                    cpContact.InReplyTo = contact.InReplyTo;
                }
                catch { }
                try
                {
                    cpContact.JobTitle = contact.JobTitle;
                }
                catch { }
                try
                {
                    cpContact.Manager = contact.Manager;
                }
                catch { }
                try
                {
                    cpContact.MiddleName = contact.MiddleName;
                }
                catch { }
                try
                {
                    cpContact.Mileage = contact.Mileage;
                }
                catch { }
                try
                {
                    cpContact.NickName = contact.NickName;
                }
                catch { }
                try
                {
                    cpContact.OfficeLocation = contact.OfficeLocation;
                }
                catch { }

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

                try
                {
                    cpContact.Profession = contact.Profession;
                }
                catch { }
                try
                {
                    cpContact.SpouseName = contact.SpouseName;
                }
                catch { }
                try
                {
                    cpContact.Subject = contact.Subject;
                }
                catch { }
                try
                {
                    cpContact.Surname = contact.Surname;
                }
                catch { }


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
                contact.Load();
                this.WriteLog("Contact Item:" + contact.Subject + ": Item Size: " + EwsHelper.CovertByteToMB(contact.Size) + ": Idem Id: " + contact.Id.ToString() + "-Error:" + erro.Message, targetMailbox);
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
                    //GC.Collect();
                }
                catch { }
            }
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
                            string email = helper.ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
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
                            string email = helper.ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
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
                            string email = helper.ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
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
                            string email = helper.ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
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
                            string email = helper.ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
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
                            string email = helper.ConvertLegacyExchangeDNToSmtpAddress(s, tgtService);
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

        #endregion

        private FolderId GetFolderId(ExchangeService svc, WellKnownFolderName searchRootFolder, string folderName)
        {
            FolderId idFolder = null;
            Folder folder = Folder.Bind(svc, searchRootFolder);
            folder.Load();

            foreach (Folder obj in folder.FindFolders(new FolderView(int.MaxValue)))
            {
                if (string.Equals(obj.DisplayName, folderName, StringComparison.CurrentCultureIgnoreCase))
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
                if (string.Equals(obj.DisplayName, folderName, StringComparison.CurrentCultureIgnoreCase))
                {
                    idFolder = obj.Id;
                }
            }

            return idFolder;
        }

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

            decimal totalItemSizeInMB = EwsHelper.CovertByteToMB(totalItemSize);

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

        private List<ItemIdentity> GetFolderIdsMigratedItems(ExchangeService svc, FolderId rootFolder)
        {
            //Metodo retorna o ID antigo da mensagem para fazer de/para na migracao
            List<ItemIdentity> itemResults = new List<ItemIdentity>();

            ItemView itemView = new ItemView(int.MaxValue);
            //Adiciona na busca a ExtendedProperty onde fica armazenado o antigo ItemID para não duplicar objetos
            itemView.PropertySet = new PropertySet(this.GetItemIDExtendedProperty());
            //Como as Throlling Policies do Exchange limitam a Find de 1000 itens
            //Precisamos fazer um workaround no código ele continua a busca até terminar todos os itens
            //Por isso o uso do Do {}while()
            FindItemsResults<Item> folderitens = null;

            do
            {
                folderitens = svc.FindItems(rootFolder, itemView);

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


            return itemResults;
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

        private string GetMailboxLanguage(ExchangeService exSrcServer, string tgtUserMail)
        {
            string language = string.Empty;
            Folder inboxFolder = Folder.Bind(exSrcServer, WellKnownFolderName.Inbox);
            string inboxFolderName = inboxFolder.DisplayName;

            ImapMigrationTool.RunSpace.RunSpaceManager runSpace = new ImapMigrationTool.RunSpace.RunSpaceManager();

            if (inboxFolderName.Equals("Inbox", StringComparison.CurrentCultureIgnoreCase))
            {
                language = "en-US";
            }
            else if (inboxFolderName.Equals("Caixa de Entrada", StringComparison.CurrentCultureIgnoreCase))
            {
                language = "pt-BR";
            }
            return language;
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
