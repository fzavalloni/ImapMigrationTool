using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;

namespace ImapMigrationTool.RunSpace
{
    public class RunSpaceManager
    {
        private RunspaceManager runspaceManager;

        public RunSpaceManager()
        {
            this.runspaceManager = new RunspaceManager();
        }

        public void SetMailboxFolderLanguages(string mailbox, string language)
        {
            PSCommand command = new PSCommand();

            command.AddCommand("Set-MailboxRegionalConfiguration");
            command.AddParameter("identity", mailbox);

            if (language == "pt-BR")
            {
                command.AddParameter("Language", language);
                command.AddParameter("DateFormat", "dd/MM/yyyy");
            }
            else
            {
                command.AddParameter("Language", "en-US");
                command.AddParameter("DateFormat", "MM/dd/yyyy");
            }

            command.AddParameter("TimeFormat", "H:mm");
            command.AddParameter("LocalizeDefaultFolderName", true);
            command.AddParameter("TimeZone", "E. South America Standard Time");

            PsExecutor(command);
        }


        private void PsExecutor(PSCommand command)
        {

            //Método que executa os comandos PowerShell, recebe apenas o objeto command
            using (Runspace runspace = runspaceManager.GetRunSpace())
            {
                runspace.Open();
                PowerShell powershell = PowerShell.Create();
                powershell.Commands = command;
                powershell.Runspace = runspace;

                try
                {
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        foreach (ErrorRecord erro in powershell.Streams.Error)
                        {
                            throw new Exception(erro.Exception.Message);
                        }
                    }
                }

                catch (Exception erro)
                {
                    throw new Exception(erro.Message);
                }

                finally
                {
                    powershell.Dispose();
                }
            }
        }

        private Collection<PSObject> GetPsExecCollecion(PSCommand command)
        {
            //Executa o comando Powershell só que retorna o resultado em uma collection
            using (Runspace runspace = runspaceManager.GetRunSpace())
            {
                runspace.Open();
                PowerShell powershell = PowerShell.Create();
                powershell.Commands = command;
                powershell.Runspace = runspace;

                try
                {
                    Collection<PSObject> results = powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        foreach (ErrorRecord erro in powershell.Streams.Error)
                        {
                            throw new Exception(erro.Exception.Message);
                        }
                    }
                    return results;
                }

                catch (Exception erro)
                {
                    throw new Exception(erro.Message);
                }

                finally
                {
                    powershell.Dispose();
                }
            }
        }
    }
}
