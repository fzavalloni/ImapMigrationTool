using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace ImapMigrationTool.RunSpace
{
    public class RunspaceManager
    {
        private string URL = "http://schemas.microsoft.com/powershell/Microsoft.Exchange";
        private string liveIdconnectionUri = string.Format("http://{0}/powershell", Environment.MachineName);
        private WSManConnectionInfo connectionInfo;

        public Runspace GetRunSpace()
        {
            this.connectionInfo = new WSManConnectionInfo(new Uri(liveIdconnectionUri), URL, PSCredential.Empty);
            this.connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Default;

            return RunspaceFactory.CreateRunspace(connectionInfo);
        }
    }
}
