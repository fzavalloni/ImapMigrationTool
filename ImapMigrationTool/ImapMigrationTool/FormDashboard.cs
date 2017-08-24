using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Configuration;
using System.IO;
using ImapMigrationTool.Entities;
using ImapMigrationTool.EwsManager;

namespace ImapMigrationTool
{
    delegate void DataGridViewRowDelegate(DataGridViewRow row);

    public partial class FormDashboard : Form
    {
        //Monta semaforo para limite de threads
        static Semaphore semaphore = new Semaphore(ThreadData.MaxThreads, 10000);
        private Dictionary<DataGridViewRow, Thread> threadList;
        List<SourceMailboxData> srcMailboxData = new List<SourceMailboxData>(); //Lista onde armazeno as senhas do Source Mailbox

        public FormDashboard()
        {
            InitializeComponent();
            this.threadList = new Dictionary<DataGridViewRow, Thread>();
        }

        private void addMailboxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            formAddMailbox f = new formAddMailbox(this);
            f.ShowDialog();
        }

        private void optionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            formOptions f = new formOptions();
            f.ShowDialog();
        }

        private int CountSemiColon(string value)
        {
            int count = 0;

            foreach (char c in value)
            {
                if (c == ';')
                {
                    count++;
                }
            }
            return count;
        }

        bool isAdminMigration = true;

        public void AddMbxToGridView(RichTextBox text)
        {
            foreach (string obj in text.Lines)
            {
                if (obj != "")
                {
                    if (obj.Contains(";"))
                    {
                        string[] array = obj.Split(';');
                        string[] row = null;

                        //Monta Grid com senha de SourcePassword
                        if (CountSemiColon(obj) == 2)
                        {
                            if (!dtGridMbx.Columns["Password"].Visible)
                            {
                                isAdminMigration = false;
                                dtGridMbx.Columns["Password"].Visible = true;
                                DataGridViewTextBoxCell TextBoxCell = new DataGridViewTextBoxCell();
                            }
                            //Monta List com senhas de usuário
                            //Isso é para nao mostrar a senha da GridView em clear text
                            //Não achei a opção de colocar PasswordChar igual ao TextBox
                            srcMailboxData.Add(new SourceMailboxData
                            {
                                Mailbox = array[0].Trim(),
                                Password = array[1].Trim()
                            });

                            //Conta caracteres da senha
                            int countLenghtPassword = array[1].Trim().Length;
                            StringBuilder sb = new StringBuilder();
                            for (int i = 0; i <= countLenghtPassword; i++)
                            {  
                                //Monta a string baseado na quantidade de caracteres da senha
                                sb.Append("*");
                            }

                            //Monsta DataGridView sem a senha
                            row = new string[] { array[0].Trim(), sb.ToString(), array[2].Trim() };
                        }
                        //Monta Grid com usuário e senha de Admin
                        else
                        {
                            row = new string[] { array[0].Trim(), null, array[1].Trim() };
                        }

                        //Se a migracao for feita via EWS, cria a aba de Mailbox Size
                        ConfigData data = new ConfigData();
                        if (data.EWSMessageMigration)
                        {
                            dtGridMbx.Columns["MailboxSize"].Visible = true;
                        }

                       
                        dtGridMbx.Rows.Add(row);
                    }
                }
            }
        }

        public void StartEWSMailboxSize(object rowObject)
        {
            DataGridViewRow row = (DataGridViewRow)rowObject;
            string sourceMailbox = dtGridMbx.Rows[row.Index].Cells["Source_Mailbox"].Value.ToString();
            string targetMailbox = dtGridMbx.Rows[row.Index].Cells["Target_Mailbox"].Value.ToString();

            row.DefaultCellStyle.BackColor = Color.White;
            Tools.SetRowValue(ref row, EColumns.results, "Starting");
            Tools.SetRowValue(ref row, EColumns.start_time, DateTime.Now.ToLocalTime().ToString());

            try
            {
                ConfigData data = new ConfigData();
                EwsMigrManager ewsMgr = new EwsMigrManager();
                //Aguarda Thread ser liberada
                Tools.SetRowValue(ref row, EColumns.results, "Queued");
                Tools.SetRowValue(ref row, EColumns.End_Time, string.Empty);
                semaphore.WaitOne();
                Tools.SetRowValue(ref row, EColumns.results, "Processing...");
                Tools.SetRowValue(ref row, EColumns.End_Time, string.Empty);

                if (isAdminMigration)
                {
                    ewsMgr.CallEWSMailboxSize(ref row, sourceMailbox, true, null);
                }
                else
                {
                    string sourcePassword = Tools.GetPasswordOnList(sourceMailbox, srcMailboxData);
                    ewsMgr.CallEWSMailboxSize(ref row, sourceMailbox, false, sourcePassword);                    
                }

                
                Tools.SetRowValue(ref row, EColumns.End_Time, DateTime.Now.ToLocalTime().ToString());

            }
            catch (Exception erro)
            {
                Tools.SetRowValue(ref row, EColumns.results, erro.Message);
                row.DefaultCellStyle.BackColor = Color.Red;
            }
            finally
            {
                //Libera recursos de Semaforo(controle de threads simultaneas) e remove thread
                semaphore.Release();
                this.Sys_RemoveThreadRow(ref row);
            }
        }

        public void StartEWSMigration(object rowObject)
        {
            DataGridViewRow row = (DataGridViewRow)rowObject;

            string sourceMailbox = dtGridMbx.Rows[row.Index].Cells["Source_Mailbox"].Value.ToString();
            string targetMailbox = dtGridMbx.Rows[row.Index].Cells["Target_Mailbox"].Value.ToString();

            row.DefaultCellStyle.BackColor = Color.White;
            Tools.SetRowValue(ref row, EColumns.results, "Starting");
            Tools.SetRowValue(ref row, EColumns.start_time, DateTime.Now.ToLocalTime().ToString());
            Tools.SetRowValue(ref row, EColumns.End_Time, string.Empty);
            try
            {
                ConfigData data = new ConfigData();
                EwsMigrManager ewsMgr = new EwsMigrManager();
                
                //Aguarda Thread ser liberada
                Tools.SetRowValue(ref row, EColumns.results, "Queued");
                Tools.SetRowValue(ref row, EColumns.End_Time, string.Empty);
                semaphore.WaitOne();
                Tools.SetRowValue(ref row, EColumns.results, "Processing...");
                Tools.SetRowValue(ref row, EColumns.End_Time, string.Empty);                

                if (!isAdminMigration)
                {
                    string sourcePassword = Tools.GetPasswordOnList(sourceMailbox, srcMailboxData);
                    ewsMgr.CallEWSMigration(ref row, sourceMailbox, targetMailbox, false, sourcePassword);
                }
                else
                {
                    ewsMgr.CallEWSMigration(ref row, sourceMailbox, targetMailbox, true, null);
                }

                Tools.SetRowValue(ref row, EColumns.results, "Finished");
                Tools.SetRowValue(ref row, EColumns.End_Time, DateTime.Now.ToLocalTime().ToString());
            }
            catch (Exception erro)
            {
                Tools.SetRowValue(ref row, EColumns.results, erro.Message);
                row.DefaultCellStyle.BackColor = Color.Red;
            }
            finally
            {
                //Libera recursos de Semaforo(controle de threads simultaneas) e remove thread
                semaphore.Release();
                this.Sys_RemoveThreadRow(ref row);
            }
        }


        public void StartImapSyncMigration(object rowObject)
        {
            ConfigData data = new ConfigData();
            bool logEnabled = data.LogEnabled;

            DataGridViewRow row = (DataGridViewRow)rowObject;
            using (Process proc = new Process())
            {
                int autenticationCountErrors = 0;
                string arguments = Tools.ReadXmlAttribute("imapSyncArgs");
                string sourceSSL = string.Empty;
                string targetSSL = string.Empty;
                string sourceMailbox = dtGridMbx.Rows[row.Index].Cells["Source_Mailbox"].Value.ToString();
                string targetMailbox = dtGridMbx.Rows[row.Index].Cells["Target_Mailbox"].Value.ToString();
                string sourceAdminUser = data.SourceAdminUser;

                proc.StartInfo.FileName = @"imapsync.exe";

                if (Convert.ToBoolean(data.SourceSSL))
                {
                    sourceSSL = " --ssl1";
                }
                if (Convert.ToBoolean(data.TargetSSL))
                {
                    targetSSL = " --ssl2";
                }


                if (isAdminMigration)
                {
                    //Nesse argumento ele insere a conta de autenticação a conta de administração
                    //Old String @"--buffersize 8192000 --nosyncacls --subscribe_all --exclude Contacts --exclude Contatos --exclude Tarefas --exclude Tasks --syncinternaldates --noauthmd5 --host1 " + data.SourceServer + " --authuser1 " + data.SourceAdminUser + " --user1 " + sourceMailbox + " --password1 " + data.SourceAdminPassword + " " + sourceSSL + " --port1 " + data.SourceImapPort + " --host2 " + data.TargetServer + " " + targetSSL + " --authuser2 " + data.TargetAdminUser + " --user2 " + targetMailbox + " --password2 " + data.TargetAdminPassword + "";
                    arguments = string.Format(arguments, data.SourceServer, sourceAdminUser, sourceMailbox, data.SourceAdminPassword, sourceSSL, data.SourceImapPort, data.TargetServer, targetSSL, data.TargetAdminUser, targetMailbox, data.TargetAdminPassword);
                }
                else
                {                    
                    string sourcePassword = Tools.GetPasswordOnList(sourceMailbox, srcMailboxData);

                    //Nesse argumento ele nao utiliza a conta de administrator                    
                    //Old String @"--buffersize 8192000 --nosyncacls --subscribe_all --exclude Contacts --exclude Contatos --exclude Tarefas --exclude Tasks --syncinternaldates --noauthmd5 --host1 " + data.SourceServer + " --authuser1 " + sourceMailbox + " --user1 " + sourceMailbox + " --password1 " + sourcePassword + " " + sourceSSL + " --port1 " + data.SourceImapPort + " --host2 " + data.TargetServer + " " + targetSSL + " --authuser2 " + data.TargetAdminUser + " --user2 " + targetMailbox + " --password2 " + data.TargetAdminPassword + "";                    
                    arguments = string.Format(arguments, data.SourceServer, sourceMailbox, sourceMailbox, sourcePassword, sourceSSL, data.SourceImapPort, data.TargetServer, targetSSL, data.TargetAdminUser, targetMailbox, data.TargetAdminPassword);
                }

                proc.StartInfo.Arguments = arguments;
                proc.StartInfo.ErrorDialog = false;
                proc.StartInfo.CreateNoWindow = true;
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.RedirectStandardInput = false;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;


                try
                {
                    //Seta cor da row para branco
                    row.DefaultCellStyle.BackColor = Color.White;

                    //Aguarda Thread ser liberada
                    Tools.SetRowValue(ref row, EColumns.results, "Queued");
                    Tools.SetRowValue(ref row, EColumns.End_Time, string.Empty);

                    semaphore.WaitOne();

                    //Verifica se o arquivo do ImapSync está no diretório atual
                    Tools.CheckImapSyncFile();

                    //Inicia processo em background
                    proc.Start();

                    Tools.SetRowValue(ref row, EColumns.steps, "1- Messages");
                    Tools.SetRowValue(ref row, EColumns.results, "Starting");
                    Tools.SetRowValue(ref row, EColumns.start_time, DateTime.Now.ToLocalTime().ToString());


                    //Pega saida do comando coloca a saida na grid correspondente e gera log, caso esteja habilitado
                    using (StreamReader reader = proc.StandardOutput)
                    {
                        string outputLine = string.Empty;

                        if (logEnabled)
                        {
                            Tools.WriteLog("============= Migration Started: " + DateTime.Now, targetMailbox);
                        }

                        while (!reader.EndOfStream)
                        {
                            outputLine = reader.ReadLine();
                            string timeNow = DateTime.Now.TimeOfDay.ToString().Remove(DateTime.Now.TimeOfDay.ToString().IndexOf('.'));

                            if (!string.IsNullOrEmpty(outputLine))
                            {
                                Tools.SetRowValue(ref row, EColumns.results, string.Format("{0}: {1}", timeNow, outputLine));
                            }

                            if (logEnabled)
                            {
                                Tools.WriteLog(string.Format("{0}: {1}", timeNow, outputLine), targetMailbox);
                            }
                            //Verifica se mostra erro de autenticação
                            if (outputLine.Contains("NO AUTHENTICATE failed"))
                            {
                                autenticationCountErrors++;
                            }
                        }
                    }

                    proc.WaitForExit();

                    Tools.SetRowValue(ref row, EColumns.steps, string.Empty);
                    Tools.SetRowValue(ref row, EColumns.results, string.Format("Finished - Autentication Errors:{0}", autenticationCountErrors));

                    //Se ouver erro de autenticação a row fica vermelha
                    if (autenticationCountErrors != 0)
                    {
                        row.DefaultCellStyle.BackColor = Color.Red;
                    }

                    Tools.SetRowValue(ref row, EColumns.End_Time, DateTime.Now.ToLocalTime().ToString());

                }
                catch (Exception erro)
                {
                    Tools.SetRowValue(ref row, EColumns.results, erro.Message);
                }
                finally
                {
                    //Libera recursos de Semaforo(controle de threads simultaneas) e remove thread
                    semaphore.Release();
                    this.Sys_RemoveThreadRow(ref row);
                }
            }
        }

        public void CheckImapCredential(object rowObject)
        {
            ConfigData data = new ConfigData();
            DataGridViewRow row = (DataGridViewRow)rowObject;
            row.DefaultCellStyle.BackColor = Color.White;

            try
            {

                string sourceMailbox = dtGridMbx.Rows[row.Index].Cells["Source_Mailbox"].Value.ToString();
                string adminMailbox = data.SourceAdminUser;
                string adminPassword = data.SourceAdminPassword;
                string sourceServer = data.SourceServer;
                int sourcePort = Convert.ToInt32(data.SourceImapPort);
                bool ssl = data.SourceSSL;

                Tools.SetRowValue(ref row, EColumns.steps, "Imap Credential");
                Tools.SetRowValue(ref row, EColumns.results, "Processing");

                if (!isAdminMigration)
                {
                    string sourcePassword = Tools.GetPasswordOnList(sourceMailbox, srcMailboxData);
                    ImapManager imapMgr = new ImapManager(sourceServer, sourcePort, ssl, sourceMailbox, sourcePassword);
                    imapMgr.Connect();
                }
                else
                {
                    ImapManager imapMgr = new ImapManager(sourceServer, sourcePort, ssl, adminMailbox, adminPassword);
                    imapMgr.Connect(sourceMailbox);
                }

                Tools.SetRowValue(ref row, EColumns.results, "User and password are VALID!");
            }
            catch (Exception erro)
            {
                Tools.SetRowValue(ref row, EColumns.results, erro.Message);
            }
            finally
            {
                //Libera recursos de Semaforo(controle de threads simultaneas) e remove thread
                this.Sys_RemoveThreadRow(ref row);
            }

        }

        private void Sys_RemoveThreadRow(ref DataGridViewRow row)
        {
            lock (threadList)
            {
                if (threadList.ContainsKey(row))
                {
                    threadList.Remove(row);
                }
            }
        }

        private void Act_StartSync(DataGridViewRow row)
        {
            lock (threadList)
            {
                if (threadList.ContainsKey(row))
                {
                    return;
                }
            }

            Thread newThread = null;
            ConfigData data = new ConfigData();

            if (data.EWSMessageMigration)
            {
                            
                //Realiza toda a migracao incluindo mensagens usando o EWS
                newThread = new Thread(StartEWSMigration);
                newThread.IsBackground = true;
            }
            else
            {
                //Realiza a migracao usando o ImapSync
                newThread = new Thread(StartImapSyncMigration);
                newThread.IsBackground = true;
            }

            lock (threadList)
            {
                threadList.Add(row, newThread);
            }

            newThread.Start(row);
        }

        private void Act_StartMbxSize(DataGridViewRow row)
        {
            lock (threadList)
            {
                if (threadList.ContainsKey(row))
                {
                    return;
                }
            }

            Thread newThread = null;
            ConfigData data = new ConfigData();

            if (data.EWSMessageMigration)
            {
                //Realiza toda a migracao incluindo mensagens usando o EWS
                newThread = new Thread(StartEWSMailboxSize);
                newThread.IsBackground = true;
            }
            else
            {
                //Se o EWS migration não estiver habilitado ele nao faz nada
                return;
            }

            lock (threadList)
            {
                threadList.Add(row, newThread);
            }

            newThread.Start(row);
        }



        private void Act_CheckImapCredentials(DataGridViewRow row)
        {
            lock (threadList)
            {
                if (threadList.ContainsKey(row))
                {
                    return;
                }
            }

            Thread newThread = new Thread(CheckImapCredential);
            newThread.IsBackground = true;

            lock (threadList)
            {
                threadList.Add(row, newThread);
            }

            newThread.Start(row);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void startImapSyncToolStripMenuItem_Click(object sender, EventArgs e)
        {         
            DataGridViewRowDelegate de = new DataGridViewRowDelegate(Act_StartSync);

            foreach (DataGridViewRow row in dtGridMbx.SelectedRows)
            {
                de.BeginInvoke(row, null, null);
            }
        }

        private void removeItensToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewRowDelegate de = new DataGridViewRowDelegate(RemoveItem);

            foreach (DataGridViewRow row in dtGridMbx.SelectedRows)
            {
                //RemoveItem(row);
                de.BeginInvoke(row, null, null);
            }
        }

        private void RemoveItem(DataGridViewRow row)
        {
            if (dtGridMbx.InvokeRequired)
            {
                DataGridViewRowDelegate de = new DataGridViewRowDelegate(RemoveItem);
                dtGridMbx.Invoke(de, row);
            }
            else
            {
                lock (row)
                {
                    lock (threadList)
                    {
                        if (threadList.ContainsKey(row))
                        {
                            threadList[row].Abort(null);
                            threadList.Remove(row);
                        }
                    }

                    dtGridMbx.Rows.Remove(row);
                }
            }
        }

        private void dtGridMbx_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            this.ShowStatusCount();
        }

        private void dtGridMbx_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            this.ShowStatusCount();
        }

        private void ShowStatusCount()
        {
            toolStripStatusLabel1.Text = string.Format("Rows: {0}", dtGridMbx.RowCount);
        }

        private void ajudaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(@".\ImapMigrationTool.chm"))
            {
                Help.ShowHelp(this, "ImapMigrationTool.chm", HelpNavigator.TableOfContents);
            }
            else
            {
                MessageBox.Show("Help file not found");
            }
        }

        private void validateImapCredentialsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewRowDelegate de = new DataGridViewRowDelegate(Act_CheckImapCredentials);

            foreach (DataGridViewRow row in dtGridMbx.SelectedRows)
            {
                de.BeginInvoke(row, null, null);
            }
        }

        private void FormDashboard_Load(object sender, EventArgs e)
        {
            int imapsyncProcsCount = Tools.GetImapsyncProcessCount();

            if (imapsyncProcsCount != 0)
            {
                DialogResult dialogResult = MessageBox.Show(string.Format("It were detected {0} imapsync.exe zumbi processes.\nWould you like to kill them?", imapsyncProcsCount), "Zumbi imapsync processes", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                {
                    Tools.KillImapsyncProcess();
                }
            }
        }

        private void getEWSMailboxSizeToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            DataGridViewRowDelegate de = new DataGridViewRowDelegate(Act_StartMbxSize);

            foreach (DataGridViewRow row in dtGridMbx.SelectedRows)
            {
                de.BeginInvoke(row, null, null);
            }
        }

        private void gcTimer_Tick(object sender, EventArgs e)
        {
            //a cada 5 minutos roda o GC
            GC.Collect();
        }

   
    }
}

public class SourceMailboxData
{
    public string Mailbox { get; set; }
    public string Password { get; set; }
}
