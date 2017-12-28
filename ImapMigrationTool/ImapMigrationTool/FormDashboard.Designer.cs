namespace ImapMigrationTool
{
    partial class FormDashboard
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.addMailboxToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sobreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ajudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dtGridMbx = new System.Windows.Forms.DataGridView();
            this.Source_Mailbox = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Password = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Target_Mailbox = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.start_time = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Steps = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.results = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MailboxSize = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.End_Time = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.getEWSMailboxSizeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.startImapSyncToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.validateImapCredentialsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.removeItensToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.status = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.gcTimer = new System.Windows.Forms.Timer(this.components);
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtGridMbx)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.status.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuToolStripMenuItem,
            this.sobreToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(936, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // menuToolStripMenuItem
            // 
            this.menuToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addMailboxToolStripMenuItem,
            this.optionsToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.menuToolStripMenuItem.Name = "menuToolStripMenuItem";
            this.menuToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.menuToolStripMenuItem.Text = "Menu";
            // 
            // addMailboxToolStripMenuItem
            // 
            this.addMailboxToolStripMenuItem.Name = "addMailboxToolStripMenuItem";
            this.addMailboxToolStripMenuItem.Size = new System.Drawing.Size(141, 22);
            this.addMailboxToolStripMenuItem.Text = "Add Mailbox";
            this.addMailboxToolStripMenuItem.Click += new System.EventHandler(this.addMailboxToolStripMenuItem_Click);
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(141, 22);
            this.optionsToolStripMenuItem.Text = "Options";
            this.optionsToolStripMenuItem.Click += new System.EventHandler(this.optionsToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(141, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // sobreToolStripMenuItem
            // 
            this.sobreToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ajudaToolStripMenuItem});
            this.sobreToolStripMenuItem.Name = "sobreToolStripMenuItem";
            this.sobreToolStripMenuItem.Size = new System.Drawing.Size(49, 20);
            this.sobreToolStripMenuItem.Text = "Sobre";
            // 
            // ajudaToolStripMenuItem
            // 
            this.ajudaToolStripMenuItem.Name = "ajudaToolStripMenuItem";
            this.ajudaToolStripMenuItem.Size = new System.Drawing.Size(105, 22);
            this.ajudaToolStripMenuItem.Text = "Ajuda";
            this.ajudaToolStripMenuItem.Click += new System.EventHandler(this.ajudaToolStripMenuItem_Click);
            // 
            // dtGridMbx
            // 
            this.dtGridMbx.AllowUserToAddRows = false;
            this.dtGridMbx.AllowUserToDeleteRows = false;
            this.dtGridMbx.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtGridMbx.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dtGridMbx.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtGridMbx.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Source_Mailbox,
            this.Password,
            this.Target_Mailbox,
            this.start_time,
            this.Steps,
            this.results,
            this.MailboxSize,
            this.End_Time});
            this.dtGridMbx.ContextMenuStrip = this.contextMenuStrip1;
            this.dtGridMbx.Cursor = System.Windows.Forms.Cursors.Arrow;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dtGridMbx.DefaultCellStyle = dataGridViewCellStyle2;
            this.dtGridMbx.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dtGridMbx.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dtGridMbx.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.dtGridMbx.Location = new System.Drawing.Point(0, 24);
            this.dtGridMbx.Name = "dtGridMbx";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtGridMbx.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dtGridMbx.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dtGridMbx.Size = new System.Drawing.Size(936, 378);
            this.dtGridMbx.TabIndex = 1;
            this.dtGridMbx.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dtGridMbx_RowsAdded);
            this.dtGridMbx.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(this.dtGridMbx_RowsRemoved);
            // 
            // Source_Mailbox
            // 
            this.Source_Mailbox.HeaderText = "Source Mailbox";
            this.Source_Mailbox.Name = "Source_Mailbox";
            this.Source_Mailbox.Width = 250;
            // 
            // Password
            // 
            this.Password.HeaderText = "Source Password";
            this.Password.Name = "Password";
            this.Password.Visible = false;
            // 
            // Target_Mailbox
            // 
            this.Target_Mailbox.HeaderText = "Target Mailbox";
            this.Target_Mailbox.Name = "Target_Mailbox";
            this.Target_Mailbox.Width = 250;
            // 
            // start_time
            // 
            this.start_time.HeaderText = "Start Time";
            this.start_time.Name = "start_time";
            this.start_time.Width = 130;
            // 
            // Steps
            // 
            this.Steps.HeaderText = "Steps";
            this.Steps.Name = "Steps";
            // 
            // results
            // 
            this.results.HeaderText = "Operation Results";
            this.results.Name = "results";
            this.results.Width = 400;
            // 
            // MailboxSize
            // 
            this.MailboxSize.HeaderText = "Mailbox Size";
            this.MailboxSize.Name = "MailboxSize";
            this.MailboxSize.Visible = false;
            this.MailboxSize.Width = 120;
            // 
            // End_Time
            // 
            this.End_Time.HeaderText = "End Time";
            this.End_Time.Name = "End_Time";
            this.End_Time.Width = 130;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.getEWSMailboxSizeToolStripMenuItem,
            this.startImapSyncToolStripMenuItem,
            this.validateImapCredentialsToolStripMenuItem,
            this.removeItensToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(225, 92);
            // 
            // getEWSMailboxSizeToolStripMenuItem
            // 
            this.getEWSMailboxSizeToolStripMenuItem.Name = "getEWSMailboxSizeToolStripMenuItem";
            this.getEWSMailboxSizeToolStripMenuItem.Size = new System.Drawing.Size(224, 22);
            this.getEWSMailboxSizeToolStripMenuItem.Text = "1- Get EWS Mailbox Size";
            this.getEWSMailboxSizeToolStripMenuItem.Click += new System.EventHandler(this.getEWSMailboxSizeToolStripMenuItem_Click);
            // 
            // startImapSyncToolStripMenuItem
            // 
            this.startImapSyncToolStripMenuItem.Name = "startImapSyncToolStripMenuItem";
            this.startImapSyncToolStripMenuItem.Size = new System.Drawing.Size(224, 22);
            this.startImapSyncToolStripMenuItem.Text = "2- Start Migration";
            this.startImapSyncToolStripMenuItem.Click += new System.EventHandler(this.startImapSyncToolStripMenuItem_Click);
            // 
            // validateImapCredentialsToolStripMenuItem
            // 
            this.validateImapCredentialsToolStripMenuItem.Name = "validateImapCredentialsToolStripMenuItem";
            this.validateImapCredentialsToolStripMenuItem.Size = new System.Drawing.Size(224, 22);
            this.validateImapCredentialsToolStripMenuItem.Text = "3 - Validate Imap Credentials";
            this.validateImapCredentialsToolStripMenuItem.Click += new System.EventHandler(this.validateImapCredentialsToolStripMenuItem_Click);
            // 
            // removeItensToolStripMenuItem
            // 
            this.removeItensToolStripMenuItem.Name = "removeItensToolStripMenuItem";
            this.removeItensToolStripMenuItem.Size = new System.Drawing.Size(224, 22);
            this.removeItensToolStripMenuItem.Text = "4 - Remove Itens";
            this.removeItensToolStripMenuItem.Click += new System.EventHandler(this.removeItensToolStripMenuItem_Click);
            // 
            // status
            // 
            this.status.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.status.Location = new System.Drawing.Point(0, 402);
            this.status.Name = "status";
            this.status.Size = new System.Drawing.Size(936, 22);
            this.status.TabIndex = 2;
            this.status.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(47, 17);
            this.toolStripStatusLabel1.Text = "Rows: 0";
            // 
            // gcTimer
            // 
            this.gcTimer.Enabled = true;
            this.gcTimer.Interval = 300000;
            this.gcTimer.Tick += new System.EventHandler(this.gcTimer_Tick);
            // 
            // FormDashboard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(936, 424);
            this.Controls.Add(this.dtGridMbx);
            this.Controls.Add(this.status);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FormDashboard";
            this.Text = "DashBoard - ImapMigrationTool- v1.6.8";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.FormDashboard_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtGridMbx)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.status.ResumeLayout(false);
            this.status.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem addMailboxToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem startImapSyncToolStripMenuItem;
        private System.Windows.Forms.DataGridView dtGridMbx;
        private System.Windows.Forms.ToolStripMenuItem removeItensToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.StatusStrip status;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripMenuItem sobreToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ajudaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem validateImapCredentialsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem getEWSMailboxSizeToolStripMenuItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn Source_Mailbox;
        private System.Windows.Forms.DataGridViewTextBoxColumn Password;
        private System.Windows.Forms.DataGridViewTextBoxColumn Target_Mailbox;
        private System.Windows.Forms.DataGridViewTextBoxColumn start_time;
        private System.Windows.Forms.DataGridViewTextBoxColumn Steps;
        private System.Windows.Forms.DataGridViewTextBoxColumn results;
        private System.Windows.Forms.DataGridViewTextBoxColumn MailboxSize;
        private System.Windows.Forms.DataGridViewTextBoxColumn End_Time;
        public System.Windows.Forms.Timer gcTimer;

    }
}

