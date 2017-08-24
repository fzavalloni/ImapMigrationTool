namespace ImapMigrationTool
{
    partial class formOptions
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Log = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblSrcPort = new System.Windows.Forms.Label();
            this.chkSrcSSL = new System.Windows.Forms.CheckBox();
            this.txtSrcImap = new System.Windows.Forms.TextBox();
            this.llbSrcPasswd = new System.Windows.Forms.Label();
            this.txtSrcAdminPass = new System.Windows.Forms.TextBox();
            this.lblSrcAdmin = new System.Windows.Forms.Label();
            this.lblSrcServerName = new System.Windows.Forms.Label();
            this.txtSrcAdminUser = new System.Windows.Forms.TextBox();
            this.txtSrcServerName = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lblTgtImapPort = new System.Windows.Forms.Label();
            this.chkTgtSSL = new System.Windows.Forms.CheckBox();
            this.txtTgtImapPort = new System.Windows.Forms.TextBox();
            this.txtTgtAdminPasswd = new System.Windows.Forms.TextBox();
            this.lblTgtAdminPasswd = new System.Windows.Forms.Label();
            this.txtTgtAdminUser = new System.Windows.Forms.TextBox();
            this.lblTgtAdminUser = new System.Windows.Forms.Label();
            this.lblTgtServerName = new System.Windows.Forms.Label();
            this.txtTgtServerName = new System.Windows.Forms.TextBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.lblLog = new System.Windows.Forms.Label();
            this.txtLogPath = new System.Windows.Forms.TextBox();
            this.chkLog = new System.Windows.Forms.CheckBox();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.lblmaxThreads = new System.Windows.Forms.Label();
            this.txtmaxThreads = new System.Windows.Forms.TextBox();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.chkBoxInboxRules = new System.Windows.Forms.CheckBox();
            this.chkBoxContacts = new System.Windows.Forms.CheckBox();
            this.chkTasks = new System.Windows.Forms.CheckBox();
            this.chkBoxCalendar = new System.Windows.Forms.CheckBox();
            this.chkBoxMessage = new System.Windows.Forms.CheckBox();
            this.chkBoxUserSetting = new System.Windows.Forms.CheckBox();
            this.chkBoxEWSMig = new System.Windows.Forms.CheckBox();
            this.lblTgtServer = new System.Windows.Forms.Label();
            this.lblSrcServer = new System.Windows.Forms.Label();
            this.cBoxTgtServer = new System.Windows.Forms.ComboBox();
            this.cBoxSrvServer = new System.Windows.Forms.ComboBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.chkBoxRemoveDuplicates = new System.Windows.Forms.CheckBox();
            this.Log.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.tabPage5.SuspendLayout();
            this.SuspendLayout();
            // 
            // Log
            // 
            this.Log.Controls.Add(this.tabPage1);
            this.Log.Controls.Add(this.tabPage2);
            this.Log.Controls.Add(this.tabPage3);
            this.Log.Controls.Add(this.tabPage4);
            this.Log.Controls.Add(this.tabPage5);
            this.Log.Location = new System.Drawing.Point(5, 7);
            this.Log.Name = "Log";
            this.Log.SelectedIndex = 0;
            this.Log.Size = new System.Drawing.Size(380, 194);
            this.Log.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.llbSrcPasswd);
            this.tabPage1.Controls.Add(this.txtSrcAdminPass);
            this.tabPage1.Controls.Add(this.lblSrcAdmin);
            this.tabPage1.Controls.Add(this.lblSrcServerName);
            this.tabPage1.Controls.Add(this.txtSrcAdminUser);
            this.tabPage1.Controls.Add(this.txtSrcServerName);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(372, 168);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Source Server";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblSrcPort);
            this.groupBox1.Controls.Add(this.chkSrcSSL);
            this.groupBox1.Controls.Add(this.txtSrcImap);
            this.groupBox1.Location = new System.Drawing.Point(207, 11);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(121, 86);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "ImapSync Settings";
            // 
            // lblSrcPort
            // 
            this.lblSrcPort.AutoSize = true;
            this.lblSrcPort.Location = new System.Drawing.Point(3, 38);
            this.lblSrcPort.Name = "lblSrcPort";
            this.lblSrcPort.Size = new System.Drawing.Size(26, 13);
            this.lblSrcPort.TabIndex = 7;
            this.lblSrcPort.Text = "Port";
            this.lblSrcPort.Click += new System.EventHandler(this.lblSrcPort_Click);
            // 
            // chkSrcSSL
            // 
            this.chkSrcSSL.AutoSize = true;
            this.chkSrcSSL.Location = new System.Drawing.Point(6, 63);
            this.chkSrcSSL.Name = "chkSrcSSL";
            this.chkSrcSSL.Size = new System.Drawing.Size(46, 17);
            this.chkSrcSSL.TabIndex = 8;
            this.chkSrcSSL.Text = "SSL";
            this.chkSrcSSL.UseVisualStyleBackColor = true;
            this.chkSrcSSL.CheckedChanged += new System.EventHandler(this.chkSrcSSL_CheckedChanged);
            // 
            // txtSrcImap
            // 
            this.txtSrcImap.Location = new System.Drawing.Point(37, 35);
            this.txtSrcImap.Name = "txtSrcImap";
            this.txtSrcImap.Size = new System.Drawing.Size(37, 20);
            this.txtSrcImap.TabIndex = 6;
            this.txtSrcImap.TextChanged += new System.EventHandler(this.txtSrcImap_TextChanged);
            // 
            // llbSrcPasswd
            // 
            this.llbSrcPasswd.AutoSize = true;
            this.llbSrcPasswd.Location = new System.Drawing.Point(6, 84);
            this.llbSrcPasswd.Name = "llbSrcPasswd";
            this.llbSrcPasswd.Size = new System.Drawing.Size(85, 13);
            this.llbSrcPasswd.TabIndex = 5;
            this.llbSrcPasswd.Text = "Admin Password";
            // 
            // txtSrcAdminPass
            // 
            this.txtSrcAdminPass.Location = new System.Drawing.Point(6, 100);
            this.txtSrcAdminPass.Name = "txtSrcAdminPass";
            this.txtSrcAdminPass.Size = new System.Drawing.Size(170, 20);
            this.txtSrcAdminPass.TabIndex = 4;
            this.txtSrcAdminPass.UseSystemPasswordChar = true;
            // 
            // lblSrcAdmin
            // 
            this.lblSrcAdmin.AutoSize = true;
            this.lblSrcAdmin.Location = new System.Drawing.Point(7, 45);
            this.lblSrcAdmin.Name = "lblSrcAdmin";
            this.lblSrcAdmin.Size = new System.Drawing.Size(87, 13);
            this.lblSrcAdmin.TabIndex = 3;
            this.lblSrcAdmin.Text = "Admin Username";
            // 
            // lblSrcServerName
            // 
            this.lblSrcServerName.AutoSize = true;
            this.lblSrcServerName.Location = new System.Drawing.Point(7, 6);
            this.lblSrcServerName.Name = "lblSrcServerName";
            this.lblSrcServerName.Size = new System.Drawing.Size(75, 13);
            this.lblSrcServerName.TabIndex = 2;
            this.lblSrcServerName.Text = "Source Server";
            // 
            // txtSrcAdminUser
            // 
            this.txtSrcAdminUser.Location = new System.Drawing.Point(6, 61);
            this.txtSrcAdminUser.Name = "txtSrcAdminUser";
            this.txtSrcAdminUser.Size = new System.Drawing.Size(170, 20);
            this.txtSrcAdminUser.TabIndex = 1;
            // 
            // txtSrcServerName
            // 
            this.txtSrcServerName.Location = new System.Drawing.Point(6, 22);
            this.txtSrcServerName.Name = "txtSrcServerName";
            this.txtSrcServerName.Size = new System.Drawing.Size(170, 20);
            this.txtSrcServerName.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.txtTgtAdminPasswd);
            this.tabPage2.Controls.Add(this.lblTgtAdminPasswd);
            this.tabPage2.Controls.Add(this.txtTgtAdminUser);
            this.tabPage2.Controls.Add(this.lblTgtAdminUser);
            this.tabPage2.Controls.Add(this.lblTgtServerName);
            this.tabPage2.Controls.Add(this.txtTgtServerName);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(372, 168);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Target Server";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lblTgtImapPort);
            this.groupBox2.Controls.Add(this.chkTgtSSL);
            this.groupBox2.Controls.Add(this.txtTgtImapPort);
            this.groupBox2.Location = new System.Drawing.Point(206, 13);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(122, 85);
            this.groupBox2.TabIndex = 9;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "ImapSync Settings";
            // 
            // lblTgtImapPort
            // 
            this.lblTgtImapPort.AutoSize = true;
            this.lblTgtImapPort.Location = new System.Drawing.Point(5, 35);
            this.lblTgtImapPort.Name = "lblTgtImapPort";
            this.lblTgtImapPort.Size = new System.Drawing.Size(26, 13);
            this.lblTgtImapPort.TabIndex = 6;
            this.lblTgtImapPort.Text = "Port";
            // 
            // chkTgtSSL
            // 
            this.chkTgtSSL.AutoSize = true;
            this.chkTgtSSL.Location = new System.Drawing.Point(8, 58);
            this.chkTgtSSL.Name = "chkTgtSSL";
            this.chkTgtSSL.Size = new System.Drawing.Size(46, 17);
            this.chkTgtSSL.TabIndex = 8;
            this.chkTgtSSL.Text = "SSL";
            this.chkTgtSSL.UseVisualStyleBackColor = true;
            // 
            // txtTgtImapPort
            // 
            this.txtTgtImapPort.Location = new System.Drawing.Point(37, 32);
            this.txtTgtImapPort.Name = "txtTgtImapPort";
            this.txtTgtImapPort.Size = new System.Drawing.Size(40, 20);
            this.txtTgtImapPort.TabIndex = 7;
            // 
            // txtTgtAdminPasswd
            // 
            this.txtTgtAdminPasswd.Location = new System.Drawing.Point(6, 103);
            this.txtTgtAdminPasswd.Name = "txtTgtAdminPasswd";
            this.txtTgtAdminPasswd.Size = new System.Drawing.Size(172, 20);
            this.txtTgtAdminPasswd.TabIndex = 5;
            this.txtTgtAdminPasswd.UseSystemPasswordChar = true;
            // 
            // lblTgtAdminPasswd
            // 
            this.lblTgtAdminPasswd.AutoSize = true;
            this.lblTgtAdminPasswd.Location = new System.Drawing.Point(8, 85);
            this.lblTgtAdminPasswd.Name = "lblTgtAdminPasswd";
            this.lblTgtAdminPasswd.Size = new System.Drawing.Size(85, 13);
            this.lblTgtAdminPasswd.TabIndex = 4;
            this.lblTgtAdminPasswd.Text = "Admin Password";
            // 
            // txtTgtAdminUser
            // 
            this.txtTgtAdminUser.Location = new System.Drawing.Point(6, 62);
            this.txtTgtAdminUser.Name = "txtTgtAdminUser";
            this.txtTgtAdminUser.Size = new System.Drawing.Size(172, 20);
            this.txtTgtAdminUser.TabIndex = 3;
            // 
            // lblTgtAdminUser
            // 
            this.lblTgtAdminUser.AutoSize = true;
            this.lblTgtAdminUser.Location = new System.Drawing.Point(8, 45);
            this.lblTgtAdminUser.Name = "lblTgtAdminUser";
            this.lblTgtAdminUser.Size = new System.Drawing.Size(87, 13);
            this.lblTgtAdminUser.TabIndex = 2;
            this.lblTgtAdminUser.Text = "Admin Username";
            // 
            // lblTgtServerName
            // 
            this.lblTgtServerName.AutoSize = true;
            this.lblTgtServerName.Location = new System.Drawing.Point(8, 6);
            this.lblTgtServerName.Name = "lblTgtServerName";
            this.lblTgtServerName.Size = new System.Drawing.Size(72, 13);
            this.lblTgtServerName.TabIndex = 1;
            this.lblTgtServerName.Text = "Target Server";
            // 
            // txtTgtServerName
            // 
            this.txtTgtServerName.Location = new System.Drawing.Point(6, 22);
            this.txtTgtServerName.Name = "txtTgtServerName";
            this.txtTgtServerName.Size = new System.Drawing.Size(172, 20);
            this.txtTgtServerName.TabIndex = 0;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.lblLog);
            this.tabPage3.Controls.Add(this.txtLogPath);
            this.tabPage3.Controls.Add(this.chkLog);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(372, 168);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Log";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // lblLog
            // 
            this.lblLog.AutoSize = true;
            this.lblLog.Location = new System.Drawing.Point(7, 4);
            this.lblLog.Name = "lblLog";
            this.lblLog.Size = new System.Drawing.Size(74, 13);
            this.lblLog.TabIndex = 2;
            this.lblLog.Text = "Directory Path";
            // 
            // txtLogPath
            // 
            this.txtLogPath.Location = new System.Drawing.Point(7, 21);
            this.txtLogPath.Name = "txtLogPath";
            this.txtLogPath.Size = new System.Drawing.Size(197, 20);
            this.txtLogPath.TabIndex = 1;
            // 
            // chkLog
            // 
            this.chkLog.AutoSize = true;
            this.chkLog.Location = new System.Drawing.Point(7, 53);
            this.chkLog.Name = "chkLog";
            this.chkLog.Size = new System.Drawing.Size(86, 17);
            this.chkLog.TabIndex = 0;
            this.chkLog.Text = "Log Enabled";
            this.chkLog.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.lblmaxThreads);
            this.tabPage4.Controls.Add(this.txtmaxThreads);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(372, 168);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Max Threads";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // lblmaxThreads
            // 
            this.lblmaxThreads.AutoSize = true;
            this.lblmaxThreads.Location = new System.Drawing.Point(7, 11);
            this.lblmaxThreads.Name = "lblmaxThreads";
            this.lblmaxThreads.Size = new System.Drawing.Size(79, 13);
            this.lblmaxThreads.TabIndex = 1;
            this.lblmaxThreads.Text = "Max Processes";
            // 
            // txtmaxThreads
            // 
            this.txtmaxThreads.Location = new System.Drawing.Point(6, 30);
            this.txtmaxThreads.Name = "txtmaxThreads";
            this.txtmaxThreads.Size = new System.Drawing.Size(136, 20);
            this.txtmaxThreads.TabIndex = 0;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.chkBoxRemoveDuplicates);
            this.tabPage5.Controls.Add(this.chkBoxInboxRules);
            this.tabPage5.Controls.Add(this.chkBoxContacts);
            this.tabPage5.Controls.Add(this.chkTasks);
            this.tabPage5.Controls.Add(this.chkBoxCalendar);
            this.tabPage5.Controls.Add(this.chkBoxMessage);
            this.tabPage5.Controls.Add(this.chkBoxUserSetting);
            this.tabPage5.Controls.Add(this.chkBoxEWSMig);
            this.tabPage5.Controls.Add(this.lblTgtServer);
            this.tabPage5.Controls.Add(this.lblSrcServer);
            this.tabPage5.Controls.Add(this.cBoxTgtServer);
            this.tabPage5.Controls.Add(this.cBoxSrvServer);
            this.tabPage5.Location = new System.Drawing.Point(4, 22);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage5.Size = new System.Drawing.Size(372, 168);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "Exchange Server";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // chkBoxInboxRules
            // 
            this.chkBoxInboxRules.AutoSize = true;
            this.chkBoxInboxRules.Location = new System.Drawing.Point(146, 49);
            this.chkBoxInboxRules.Name = "chkBoxInboxRules";
            this.chkBoxInboxRules.Size = new System.Drawing.Size(109, 17);
            this.chkBoxInboxRules.TabIndex = 11;
            this.chkBoxInboxRules.Text = "Copy Inbox Rules";
            this.chkBoxInboxRules.UseVisualStyleBackColor = true;
            // 
            // chkBoxContacts
            // 
            this.chkBoxContacts.AutoSize = true;
            this.chkBoxContacts.Location = new System.Drawing.Point(146, 26);
            this.chkBoxContacts.Name = "chkBoxContacts";
            this.chkBoxContacts.Size = new System.Drawing.Size(95, 17);
            this.chkBoxContacts.TabIndex = 10;
            this.chkBoxContacts.Text = "Copy Contacts";
            this.chkBoxContacts.UseVisualStyleBackColor = true;
            // 
            // chkTasks
            // 
            this.chkTasks.AutoSize = true;
            this.chkTasks.Location = new System.Drawing.Point(39, 72);
            this.chkTasks.Name = "chkTasks";
            this.chkTasks.Size = new System.Drawing.Size(82, 17);
            this.chkTasks.TabIndex = 9;
            this.chkTasks.Text = "Copy Tasks";
            this.chkTasks.UseVisualStyleBackColor = true;
            // 
            // chkBoxCalendar
            // 
            this.chkBoxCalendar.AutoSize = true;
            this.chkBoxCalendar.Location = new System.Drawing.Point(39, 49);
            this.chkBoxCalendar.Name = "chkBoxCalendar";
            this.chkBoxCalendar.Size = new System.Drawing.Size(95, 17);
            this.chkBoxCalendar.TabIndex = 8;
            this.chkBoxCalendar.Text = "Copy Calendar";
            this.chkBoxCalendar.UseVisualStyleBackColor = true;
            // 
            // chkBoxMessage
            // 
            this.chkBoxMessage.AutoSize = true;
            this.chkBoxMessage.Location = new System.Drawing.Point(39, 26);
            this.chkBoxMessage.Name = "chkBoxMessage";
            this.chkBoxMessage.Size = new System.Drawing.Size(101, 17);
            this.chkBoxMessage.TabIndex = 7;
            this.chkBoxMessage.Text = "Copy Messages";
            this.chkBoxMessage.UseVisualStyleBackColor = true;
            // 
            // chkBoxUserSetting
            // 
            this.chkBoxUserSetting.AutoSize = true;
            this.chkBoxUserSetting.Location = new System.Drawing.Point(8, 99);
            this.chkBoxUserSetting.Name = "chkBoxUserSetting";
            this.chkBoxUserSetting.Size = new System.Drawing.Size(332, 17);
            this.chkBoxUserSetting.TabIndex = 6;
            this.chkBoxUserSetting.Text = "Keep Default Folder Language (Must be run in Exchange Server)";
            this.chkBoxUserSetting.UseVisualStyleBackColor = true;
            // 
            // chkBoxEWSMig
            // 
            this.chkBoxEWSMig.AutoSize = true;
            this.chkBoxEWSMig.Location = new System.Drawing.Point(7, 5);
            this.chkBoxEWSMig.Name = "chkBoxEWSMig";
            this.chkBoxEWSMig.Size = new System.Drawing.Size(167, 17);
            this.chkBoxEWSMig.TabIndex = 5;
            this.chkBoxEWSMig.Text = "Migrate messages using EWS";
            this.chkBoxEWSMig.UseVisualStyleBackColor = true;
            this.chkBoxEWSMig.CheckedChanged += new System.EventHandler(this.chkBoxEWSMig_CheckedChanged);
            // 
            // lblTgtServer
            // 
            this.lblTgtServer.AutoSize = true;
            this.lblTgtServer.Location = new System.Drawing.Point(201, 128);
            this.lblTgtServer.Name = "lblTgtServer";
            this.lblTgtServer.Size = new System.Drawing.Size(127, 13);
            this.lblTgtServer.TabIndex = 4;
            this.lblTgtServer.Text = "Target Exchange Version";
            // 
            // lblSrcServer
            // 
            this.lblSrcServer.AutoSize = true;
            this.lblSrcServer.Location = new System.Drawing.Point(0, 128);
            this.lblSrcServer.Name = "lblSrcServer";
            this.lblSrcServer.Size = new System.Drawing.Size(130, 13);
            this.lblSrcServer.TabIndex = 3;
            this.lblSrcServer.Text = "Source Exchange Version";
            // 
            // cBoxTgtServer
            // 
            this.cBoxTgtServer.FormattingEnabled = true;
            this.cBoxTgtServer.Location = new System.Drawing.Point(204, 144);
            this.cBoxTgtServer.Name = "cBoxTgtServer";
            this.cBoxTgtServer.Size = new System.Drawing.Size(162, 21);
            this.cBoxTgtServer.TabIndex = 2;
            // 
            // cBoxSrvServer
            // 
            this.cBoxSrvServer.FormattingEnabled = true;
            this.cBoxSrvServer.Location = new System.Drawing.Point(3, 144);
            this.cBoxSrvServer.Name = "cBoxSrvServer";
            this.cBoxSrvServer.Size = new System.Drawing.Size(162, 21);
            this.cBoxSrvServer.TabIndex = 1;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(5, 205);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(376, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // chkBoxRemoveDuplicates
            // 
            this.chkBoxRemoveDuplicates.AutoSize = true;
            this.chkBoxRemoveDuplicates.Location = new System.Drawing.Point(146, 73);
            this.chkBoxRemoveDuplicates.Name = "chkBoxRemoveDuplicates";
            this.chkBoxRemoveDuplicates.Size = new System.Drawing.Size(119, 17);
            this.chkBoxRemoveDuplicates.TabIndex = 12;
            this.chkBoxRemoveDuplicates.Text = "Remove Duplicates";
            this.chkBoxRemoveDuplicates.UseVisualStyleBackColor = true;
            // 
            // formOptions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(386, 231);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.Log);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "formOptions";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Options";
            this.Log.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            this.tabPage5.ResumeLayout(false);
            this.tabPage5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl Log;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.CheckBox chkSrcSSL;
        private System.Windows.Forms.Label lblSrcPort;
        private System.Windows.Forms.TextBox txtSrcImap;
        private System.Windows.Forms.Label llbSrcPasswd;
        private System.Windows.Forms.TextBox txtSrcAdminPass;
        private System.Windows.Forms.Label lblSrcAdmin;
        private System.Windows.Forms.Label lblSrcServerName;
        private System.Windows.Forms.TextBox txtSrcAdminUser;
        private System.Windows.Forms.TextBox txtSrcServerName;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.CheckBox chkTgtSSL;
        private System.Windows.Forms.TextBox txtTgtImapPort;
        private System.Windows.Forms.Label lblTgtImapPort;
        private System.Windows.Forms.TextBox txtTgtAdminPasswd;
        private System.Windows.Forms.Label lblTgtAdminPasswd;
        private System.Windows.Forms.TextBox txtTgtAdminUser;
        private System.Windows.Forms.Label lblTgtAdminUser;
        private System.Windows.Forms.Label lblTgtServerName;
        private System.Windows.Forms.TextBox txtTgtServerName;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Label lblLog;
        private System.Windows.Forms.TextBox txtLogPath;
        private System.Windows.Forms.CheckBox chkLog;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Label lblmaxThreads;
        private System.Windows.Forms.TextBox txtmaxThreads;
        private System.Windows.Forms.TabPage tabPage5;
        private System.Windows.Forms.Label lblTgtServer;
        private System.Windows.Forms.Label lblSrcServer;
        private System.Windows.Forms.ComboBox cBoxTgtServer;
        private System.Windows.Forms.ComboBox cBoxSrvServer;
        private System.Windows.Forms.CheckBox chkBoxEWSMig;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox chkBoxUserSetting;
        private System.Windows.Forms.CheckBox chkBoxContacts;
        private System.Windows.Forms.CheckBox chkTasks;
        private System.Windows.Forms.CheckBox chkBoxCalendar;
        private System.Windows.Forms.CheckBox chkBoxMessage;
        private System.Windows.Forms.CheckBox chkBoxInboxRules;
        private System.Windows.Forms.CheckBox chkBoxRemoveDuplicates;
    }
}