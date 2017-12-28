using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace ImapMigrationTool.EwsManager
{
    public class EwsFolderMgr
    {
        public void CreateNewFolder(ExchangeService svc, string folderName, FolderId folderId, string folderClass)
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

        public void CreateNewFolder(ExchangeService svc, string folderName, WellKnownFolderName folderWellKnow, string folderClass)
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
                ValidationCreationFolderException(erro);
            }
        }

        private void ValidationCreationFolderException(Exception erro)
        {
            if (!erro.Message.Contains("A folder with the specified name already exists."))
            {
                throw new Exception(erro.Message);
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
                ValidationCreationFolderException(erro);
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
                ValidationCreationFolderException(erro);
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
                ValidationCreationFolderException(erro);
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
                ValidationCreationFolderException(erro);
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
                ValidationCreationFolderException(erro);
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
                ValidationCreationFolderException(erro);
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
                ValidationCreationFolderException(erro);
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
                ValidationCreationFolderException(erro);
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
                ValidationCreationFolderException(erro);
            }
        }

        #endregion

    }
}
