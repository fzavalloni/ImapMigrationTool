namespace ImapMigrationTool
{
    partial class formAddMailbox
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formAddMailbox));
            this.mbxText = new System.Windows.Forms.RichTextBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblLines = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // mbxText
            // 
            this.mbxText.Location = new System.Drawing.Point(3, 11);
            this.mbxText.Name = "mbxText";
            this.mbxText.Size = new System.Drawing.Size(530, 269);
            this.mbxText.TabIndex = 0;
            this.mbxText.Text = "";
            this.mbxText.TextChanged += new System.EventHandler(this.mbxText_TextChanged);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(3, 302);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(530, 27);
            this.btnAdd.TabIndex = 1;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(101, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(331, 156);
            this.label1.TabIndex = 2;
            this.label1.Text = resources.GetString("label1.Text");
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Gold;
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(3, 335);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(530, 173);
            this.panel1.TabIndex = 2;
            // 
            // lblLines
            // 
            this.lblLines.AutoSize = true;
            this.lblLines.Location = new System.Drawing.Point(7, 283);
            this.lblLines.Name = "lblLines";
            this.lblLines.Size = new System.Drawing.Size(44, 13);
            this.lblLines.TabIndex = 3;
            this.lblLines.Text = "Lines: 0";
            // 
            // formAddMailbox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(538, 510);
            this.Controls.Add(this.lblLines);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.mbxText);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "formAddMailbox";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Mailbox Migration List";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox mbxText;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblLines;

    }
}