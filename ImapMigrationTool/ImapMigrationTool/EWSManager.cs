using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Security.Principal;
using ImapMigrationTool.Entities;
using System.Windows.Forms;
using ImapMigrationTool.ExchangeService_2010;
namespace ImapMigrationTool
{
    public class EWSManager
    {

        private ExchangeServiceBinding GetService(string adminUserName, string adminPassword, string userMail, string ewsUrl, ExchangeVersionType exVersion)
        {            
            //Ignorar validação de certificado
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(OnValidationCallback);

            // Identify the service binding and the user.

            ExchangeServiceBinding service = new ExchangeServiceBinding();
            service.RequestServerVersionValue = new RequestServerVersion();
            service.RequestServerVersionValue.Version = exVersion;
            service.Credentials = new NetworkCredential(adminUserName, adminPassword);
            service.Url = ewsUrl;

            //Impersonating user context    
            ExchangeImpersonationType exExchangeImpersonation = new ExchangeImpersonationType();
            ConnectingSIDType csConnectingSid = new ConnectingSIDType();
            csConnectingSid.Item = userMail;
            exExchangeImpersonation.ConnectingSID = csConnectingSid;
            service.ExchangeImpersonation = exExchangeImpersonation;

            return service;
        }


        private static bool OnValidationCallback(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }

        public void CallEWS(ref DataGridViewRow row, string srcUserMail, string tgtUserMail)
        {

            try
            {
                ConfigData data = new ConfigData();

                string ewsSrcUrl = string.Format("https://{0}/ews/exchange.asmx", data.SourceServer);
                ExchangeServiceBinding exSrcServer = this.GetService(data.SourceAdminUser, data.SourceAdminPassword, srcUserMail, ewsSrcUrl, ExchangeVersionType.Exchange2010_SP1);

                string ewsTgtUrl = string.Format("https://{0}/ews/exchange.asmx", data.TargetServer);

                ExchangeServiceBinding exTgtServer = this.GetService(data.TargetAdminUser, data.TargetAdminPassword, tgtUserMail, ewsTgtUrl, ExchangeVersionType.Exchange2010_SP1);

                //Migrando Tarefas
                Tools.SetRowValue(ref row, EColumns.steps, "2- Tasks");
                Tools.SetRowValue(ref row, EColumns.results, string.Empty);

                MigrateItens(ref row,exSrcServer, exTgtServer,DistinguishedFolderIdNameType.contacts, srcUserMail, tgtUserMail);

                //this.MigrateTasks(ref row, exSrcServer, exTgtServer);

                ////Migrando Contatos
                //Tools.SetRowValue(ref row, EColumns.steps, "3- Contacts");
                //Tools.SetRowValue(ref row, EColumns.results, string.Empty);

                //this.MigrateContacts(ref row, exSrcServer, exTgtServer);

                ////Migrando Calendario
                //Tools.SetRowValue(ref row, EColumns.steps, "4- Calendars");
                //Tools.SetRowValue(ref row, EColumns.results, string.Empty);

                //this.MigrateCalendar(ref row, exSrcServer, exTgtServer);
            }
            catch (Exception erro)
            {
                throw new Exception("Erro:" + erro.Message);
            }

        }

        private void MigrateItens(ref DataGridViewRow row, ExchangeServiceBinding srcService, ExchangeServiceBinding tgtService, DistinguishedFolderIdNameType folder, string srcUserMail, string tgtUserMail)
        {
            ResponseMessageType[] itemsResponse = this.GetFolderItems(srcService, folder);

            ExportItemsType exExportItems = new ExportItemsType();

            foreach (ResponseMessageType responseMessage in itemsResponse)
            {
                FindItemResponseMessageType firmt = responseMessage as FindItemResponseMessageType;
                FindItemParentType fipt = firmt.RootFolder;
                object obj = fipt.Item;
                int count = 0;

                // FindItem contains an array of items.
                if (obj is ArrayOfRealItemsType)
                {
                    ArrayOfRealItemsType items = (obj as ArrayOfRealItemsType);

                    exExportItems.ItemIds = new ItemIdType[(items.Items.Count() + 1)];

                    foreach (ItemType it in items.Items)
                    {

                        exExportItems.ItemIds[count] = new ItemIdType();
                        exExportItems.ItemIds[count].Id = it.ItemId.Id;
                        count++;
                    }
                }
            }

            ExportItemsResponseType exResponse = srcService.ExportItems(exExportItems);
            ResponseMessageType[] rmResponses = exResponse.ResponseMessages.Items;
            UploadItemsType upUploadItems = new UploadItemsType();
            upUploadItems.Items = new UploadItemType[(rmResponses.Length + 1)];
            Int32 icItemCount = 0;

            foreach (ResponseMessageType rmReponse in rmResponses)
            {
                if (rmReponse.ResponseClass == ResponseClassType.Success)
                {
                    ExportItemsResponseMessageType exExportedItem = (ExportItemsResponseMessageType)rmReponse;
                    Byte[] messageBytes = exExportedItem.Data;
                    UploadItemType upUploadItem = new UploadItemType();
                    upUploadItem.CreateAction = CreateActionType.UpdateOrCreate;
                    upUploadItem.Data = messageBytes;
                    upUploadItem.IsAssociatedSpecified = true;
                    upUploadItem.IsAssociated = false;
                    upUploadItems.Items[icItemCount] = upUploadItem;

                    FolderIdManager folderIdMgr = new FolderIdManager();
                    FolderIdType folderId = new FolderIdType();
                    folderId.Id = folderIdMgr.GetFolderId(tgtUserMail, Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Contacts, Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2010_SP2);


                    upUploadItem.ParentFolderId = folderId;
                    icItemCount += 1;
                }
            }
            //Erro de Internal Server Error nessa etapa
            UploadItemsResponseType upLoadResponse = tgtService.UploadItems(upUploadItems);
            Int32 Success = 0;
            Int32 Failure = 0;
            foreach (ResponseMessageType upResponse in upLoadResponse.ResponseMessages.Items)
            {
                if (upResponse.ResponseClass == ResponseClassType.Success)
                {
                    Success++;
                }
                if (upResponse.ResponseClass == ResponseClassType.Error)
                {
                    Failure++;
                }
            }

          
            string resTask = string.Format("Items Copied Sucessfull : {0} - Failure: {1}", Success, Failure);

            Tools.SetRowValue(ref row, EColumns.results, resTask);
                
            //iv.Offset += fiItems.Items.Count;

        }

        private ResponseMessageType[] GetFolderItems(ExchangeServiceBinding svc, DistinguishedFolderIdNameType folder)
        {
            // Form the FindItem request.
            FindItemType findItemRequest = new FindItemType();
            findItemRequest.Traversal = ItemQueryTraversalType.Shallow;

            // Define which item properties are returned in the response.
            ItemResponseShapeType itemProperties = new ItemResponseShapeType();
            itemProperties.BaseShape = DefaultShapeNamesType.AllProperties;

            //Define propriedade que armazena antigo ID
            PathToExtendedFieldType netShowUrlPath = new PathToExtendedFieldType();
            netShowUrlPath.PropertyTag = "0x3A4D";
            netShowUrlPath.PropertyType = MapiPropertyTypeType.String;

            //Adiciona propriedade na busca
            itemProperties.AdditionalProperties = new BasePathToElementType[1];
            itemProperties.AdditionalProperties[0] = netShowUrlPath;

            // Add properties shape to the request.
            findItemRequest.ItemShape = itemProperties;

            // Identify which folders to search to find items.
            DistinguishedFolderIdType[] folderIDArray = new DistinguishedFolderIdType[2];
            folderIDArray[0] = new DistinguishedFolderIdType();
            folderIDArray[0].Id = folder;

            // Add folders to the request.
            findItemRequest.ParentFolderIds = folderIDArray;

            // Send the request and get the response.
            FindItemResponseType findItemResponse = svc.FindItem(findItemRequest);

            // Get the response messages.
            ResponseMessageType[] rmta = findItemResponse.ResponseMessages.Items;

            return rmta;
        }


        public FolderIdType FindFolderID(ExchangeServiceBinding service, DistinguishedFolderIdNameType folder)
        {
            
            FindFolderType requestFindFolder = new FindFolderType();
            requestFindFolder.Traversal = FolderQueryTraversalType.Deep;

            DistinguishedFolderIdType[] folderIDArray = new DistinguishedFolderIdType[1];
            folderIDArray[0] = new DistinguishedFolderIdType();
            folderIDArray[0].Id = folder;

            FolderResponseShapeType itemProperties = new FolderResponseShapeType();
            itemProperties.BaseShape = DefaultShapeNamesType.AllProperties;

            requestFindFolder.ParentFolderIds = folderIDArray;
            requestFindFolder.FolderShape = itemProperties;
            //requestFindFolder.FolderShape.BaseShape = DefaultShapeNamesType.AllProperties;
            

            FindFolderResponseType objFindFolderResponse = service.FindFolder(requestFindFolder);

            foreach (ResponseMessageType responseMsg in objFindFolderResponse.ResponseMessages.Items)
            {
                if (responseMsg.ResponseClass == ResponseClassType.Success)
                {
                    FindFolderResponseMessageType objFindResponse = responseMsg as FindFolderResponseMessageType;

                    foreach (BaseFolderType objFolderType in objFindResponse.RootFolder.Folders)
                    {
                        return objFolderType.FolderId;
                    }
                }
            }
            return null;
        }







    }




}
