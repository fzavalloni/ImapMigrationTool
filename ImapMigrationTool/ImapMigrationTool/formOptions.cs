using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ImapMigrationTool.Entities;
using Microsoft.Exchange.WebServices.Data;

namespace ImapMigrationTool
{
    public partial class formOptions : Form
    {
        public formOptions()
        {
            InitializeComponent();
            this.FillData();
        }

        private void FillData()
        {
            try
            {
                ConfigData data = new ConfigData();

                txtSrcServerName.Text = data.SourceServer;
                txtSrcAdminUser.Text = data.SourceAdminUser;
                txtSrcAdminPass.Text = data.SourceAdminPassword;
                txtSrcImap.Text = data.SourceImapPort;
                chkSrcSSL.Checked = data.SourceSSL;                

                txtTgtServerName.Text = data.TargetServer;
                txtTgtAdminUser.Text = data.TargetAdminUser;
                txtTgtAdminPasswd.Text = data.TargetAdminPassword;
                txtTgtImapPort.Text = data.TargetImapPort;
                chkTgtSSL.Checked = data.TargetSSL;

                txtLogPath.Text = data.LogPath;
                chkLog.Checked = data.LogEnabled;
                txtmaxThreads.Text = data.MaxThreads;
                //chkExMigration.Checked = data.CalendarMigration;
                chkBoxEWSMig.Checked = data.EWSMessageMigration;
                chkBoxMessage.Checked = data.EWSMessage;
                chkBoxContacts.Checked = data.EWSContacts;
                chkBoxCalendar.Checked = data.EWSCalendar;
                chkTasks.Checked =  data.EWSTasks;
                chkBoxInboxRules.Checked = data.EWSInboxRules;
                chkBoxUserSetting.Checked = data.MigrateDefaultFolderName;
                chkBoxRemoveDuplicates.Checked = data.RemoveDuplicate;

                
                //Fill Exchange Combobox
                cBoxSrvServer.DataSource = Enum.GetValues(typeof(ExchangeVersion));
                cBoxSrvServer.SelectedIndex = cBoxSrvServer.FindStringExact(data.SrcServerVersion);
                
                cBoxTgtServer.DataSource = Enum.GetValues(typeof(ExchangeVersion));
                cBoxTgtServer.SelectedIndex = cBoxTgtServer.FindStringExact(data.TgtServerVersion);

            }
            catch (Exception erro)
            {
                MessageBox.Show("Erro: " + erro.Message);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                ConfigData data = new ConfigData();

                data.SourceServer = txtSrcServerName.Text;
                data.SourceAdminUser = txtSrcAdminUser.Text;
                data.SourceAdminPassword = txtSrcAdminPass.Text;
                data.SourceImapPort = txtSrcImap.Text;
                data.SourceSSL = chkSrcSSL.Checked;                

                data.TargetServer = txtTgtServerName.Text;
                data.TargetAdminUser = txtTgtAdminUser.Text;
                data.TargetAdminPassword = txtTgtAdminPasswd.Text;
                data.TargetImapPort = txtTgtImapPort.Text;
                data.TargetSSL = chkTgtSSL.Checked;
                data.LogPath = txtLogPath.Text;
                data.LogEnabled = chkLog.Checked;
                data.MaxThreads = txtmaxThreads.Text;                
                data.EWSMessageMigration = chkBoxEWSMig.Checked;
                data.EWSMessage = chkBoxMessage.Checked;
                data.EWSContacts = chkBoxContacts.Checked;
                data.EWSCalendar = chkBoxCalendar.Checked;
                data.EWSTasks = chkTasks.Checked;
                data.EWSInboxRules = chkBoxInboxRules.Checked;
                data.RemoveDuplicate = chkBoxRemoveDuplicates.Checked;

                data.MigrateDefaultFolderName = chkBoxUserSetting.Checked;

                data.SrcServerVersion = cBoxSrvServer.SelectedItem.ToString();
                data.TgtServerVersion = cBoxTgtServer.SelectedItem.ToString();

                this.Close();
            }
            catch (Exception erro)
            {
                MessageBox.Show("Erro: " + erro.Message);
            }
        }

        private void txtSrcImap_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblSrcPort_Click(object sender, EventArgs e)
        {

        }

        private void chkSrcSSL_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkBoxEWSMig_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBoxEWSMig.Checked == true)
            {
                chkBoxMessage.Checked = true;
                chkBoxCalendar.Checked = true;
                chkBoxContacts.Checked = true;
                chkTasks.Checked = true;
                chkBoxInboxRules.Checked = true;
                chkBoxRemoveDuplicates.Checked = true;
            }
            if (chkBoxEWSMig.Checked == false)
            {
                chkBoxMessage.Checked = false;
                chkBoxCalendar.Checked = false;
                chkBoxContacts.Checked = false;
                chkTasks.Checked = false;
                chkBoxInboxRules.Checked = false;
                chkBoxRemoveDuplicates.Checked = false;
            }
        }
        
    }
}
