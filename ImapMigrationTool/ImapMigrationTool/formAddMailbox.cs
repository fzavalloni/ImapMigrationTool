using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ImapMigrationTool
{
    public partial class formAddMailbox : Form
    {
        FormDashboard formDashborad;

        public formAddMailbox()
        {
            InitializeComponent();
        }

        public formAddMailbox(FormDashboard formDashborad)
        {
            InitializeComponent();
            this.formDashborad = formDashborad;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(mbxText.Text))
            {
                if (mbxText.Lines.Length >= 10000)
                {
                    MessageBox.Show("Error: Limit lines has reached: 10000");
                }
                else
                {
                    this.formDashborad.AddMbxToGridView(mbxText);
                    this.Close();
                }
            }
        }

        private void mbxText_TextChanged(object sender, EventArgs e)
        {
            lblLines.Text = string.Format("Lines: {0}", mbxText.Lines.Length);
        }
    }
}
