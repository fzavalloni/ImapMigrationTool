using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace ImapMigrationTool.EwsManager
{
    public class EwsHelper
    {
        //Faz uma busca no Exchange pelo legacyDN e retorna o emailAddress
        public string ConvertLegacyExchangeDNToSmtpAddress(EmailAddress email, ExchangeService svc)
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

        public static decimal CovertByteToMB(Int64 value)
        {
            //Converte o tamanho de byte para MB/ O decimal.Round deixa apenas 2 casas decimais
            return decimal.Round(Convert.ToDecimal(((value / 1024f) / 1024f)), 2);
        }          

        public bool IsRuleAlreadyCreated(string ruleDisplayName, RuleCollection ruleCollection)
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

        public List<FolderIdentity> RemoveFoldersAlreadyExists(List<FolderIdentity> folderListSourceSrv, List<FolderIdentity> folderListTargetSrv)
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

        public FolderId PickCurrentObjInList(FolderIdentity obj, List<FolderIdentity> list)
        {
            FolderId id = null;

            foreach (FolderIdentity item in list)
            {
                if (string.Equals(obj.FolderPath,item.FolderPath,StringComparison.CurrentCultureIgnoreCase))
                {
                    id = item.FolderID;
                }
            }
            return id;
        }

        public int SumFolderCountMessages(List<FolderIdentity> list)
        {
            int sum = 0;

            foreach (FolderIdentity item in list)
            {
                sum = sum + item.CountItems;
            }

            return sum;
        }

        public int GetDeletedItensCount(ExchangeService svc)
        {
            Folder folder = Folder.Bind(svc, WellKnownFolderName.DeletedItems);
            return folder.TotalCount;
        }

        public Attendee GetCorrectedAttendee(Attendee attendee)
        {
            Attendee newAttendee = new Attendee();
            newAttendee.Address = attendee.Address;
            newAttendee.Name = attendee.Name;            

            return newAttendee;
        }
    }
}
