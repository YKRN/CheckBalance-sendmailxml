using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Configuration;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        ToolTip toolTip = new ToolTip();
        string richTextToolTip = "ToolTip Message Here";
        static Regex validate_emailaddress = email_validation();
        public Form1()
        {
            InitializeComponent();
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage3"];
            tabControl1.Appearance = TabAppearance.FlatButtons;
            tabControl1.ItemSize = new Size(0, 1);
            tabControl1.SizeMode = TabSizeMode.Fixed;
            ContextMenu cm = new ContextMenu();
            cm.MenuItems.Add("Export txt", new EventHandler(exportLogFile));
            cm.MenuItems.Add("Clear", new EventHandler(clearLog));
            m_dataGridViewMailList.ContextMenu = cm;
            m_txtHost.Text = config.AppSettings.Settings["Server"].Value;
            m_txtPort.Text = config.AppSettings.Settings["Port"].Value;
            m_txtUid.Text = config.AppSettings.Settings["Uid"].Value;
            m_txtPassword.Text = config.AppSettings.Settings["Password"].Value;
            m_txtDatabaseName.Text = config.AppSettings.Settings["Database"].Value;
            m_txtTimeout.Text = config.AppSettings.Settings["Timeout"].Value;
            lblPath.Text = "NONE";

            toolTip.OwnerDraw = true;
            toolTip.Draw += new DrawToolTipEventHandler(toolTip1_Draw);
            toolTip.Popup += new PopupEventHandler(toolTip1_Popup);
          


            m_txtSMTPHost.Text = config.AppSettings.Settings["SMTPHOST"].Value;
            m_txtSMTPPort.Text = config.AppSettings.Settings["SMTPPort"].Value;
            m_txtSMTPUserName.Text = config.AppSettings.Settings["SMTPUserName"].Value;
            m_txtSMTPPassword.Text = config.AppSettings.Settings["SMTPPassword"].Value;
            m_checkBoxSSL.Checked = bool.Parse(config.AppSettings.Settings["SMTPUseSSL"].Value);
            checkBox2.Checked = bool.Parse(config.AppSettings.Settings["UseLogFile"].Value);
            m_txtLogPath.Text = config.AppSettings.Settings["LogFilePath"].Value;
            m_txtcc.Text = config.AppSettings.Settings["CC"].Value;
            m_txtbcc.Text=config.AppSettings.Settings["BCC"].Value;
            //  MessageBox.Show(config.AppSettings.Settings["SMTPUseSSL"].Value);
            config.Save(ConfigurationSaveMode.Modified);
            m_dataGridViewMailList.AllowUserToAddRows = false;

        }
        private static Regex email_validation()
        {
            string pattern = @"^(?!\.)(""([^""\r\\]|\\[""\r\\])*""|"
                + @"([-a-z0-9!#$%&'*+/=?^_`{|}~]|(?<!\.)\.)*)(?<!\.)"
                + @"@[a-z0-9][\w\.-]*[a-z0-9]\.[a-z][a-z\.]*[a-z]$";

            return new Regex(pattern, RegexOptions.IgnoreCase);
        }
        void toolTip1_Popup(object sender, PopupEventArgs e)
        {

            // on popip set the size of tool tip
            e.ToolTipSize = TextRenderer.MeasureText(richTextToolTip, new Font("Arial", 16.0f));
        }

        void toolTip1_Draw(object sender, DrawToolTipEventArgs e)
        {
            Font f = new Font("Arial", 16.0f);
            e.DrawBackground();
            e.DrawBorder();
            richTextToolTip = e.ToolTipText;
            e.Graphics.DrawString(e.ToolTipText, f, Brushes.Black, new PointF(2, 2));
        }
        private void clearLog(object sender, EventArgs e)
        {
            m_txtLog.Text = "";
        }
        private void exportLogFile(object sender, EventArgs e)
        {

            //MessageBox.Show("im here");


        }
        DataGridViewRow row;
        private void button1_Click(object sender, EventArgs e)
        {

            m_dataGridViewMailList.AllowUserToAddRows = true;
            row = (DataGridViewRow)m_dataGridViewMailList.Rows[0].Clone();


            string secilendosyayolu;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
           
              
            openFileDialog1.Filter = "XML File |*.xml|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                m_dataGridViewMailList.AllowDrop = false;
                return;
            }

            // openFileDialog1.ShowDialog();
            secilendosyayolu = openFileDialog1.FileName;
        
            if (secilendosyayolu == "")
            {
                lblPath.Text = "NONE";
            }
            else
            {
           row.DefaultCellStyle.BackColor = Color.White;
                lblPath.Text = secilendosyayolu;

                XmlDocument doc = new XmlDocument();
                doc.Load(secilendosyayolu);
                XmlElement root = doc.DocumentElement;

                XmlNodeList nodesDbtr = root.GetElementsByTagName("Dbtr");
                foreach (XmlNode node in nodesDbtr)
                {

                    row.Cells[0].Value = node["Nm"].InnerText;
                    //row.Cells[2].Value = node["Id"].InnerText;
                    //MessageBox.Show(node["Nm"].InnerText);
                    //MessageBox.Show(node["Id"].InnerText);




                }

                XmlNodeList nodes = root.GetElementsByTagName("Cdtr");
                foreach (XmlNode node in nodes)
                {

                    row.Cells[1].Value = node["Nm"].InnerText;
                    row.Cells[2].Value = node["Id"].InnerText;
                    //MessageBox.Show(node["Nm"].InnerText);
                    //MessageBox.Show(node["Id"].InnerText);




                }
                XmlNodeList nodesAMT = root.GetElementsByTagName("Amt");


                foreach (XmlNode node in nodesAMT)
                {


                    // MessageBox.Show(node["InstdAmt"].InnerText);
                    row.Cells[3].Value = node["InstdAmt"].InnerText;

                }

                XmlNodeList nodesID = root.GetElementsByTagName("CdtrAcct");


                foreach (XmlNode node in nodesID)
                {


                    // MessageBox.Show(node["Id"].InnerText);
                    row.Cells[4].Value = node["Id"].InnerText;

                }

                XmlNodeList nodesBankRef = root.GetElementsByTagName("RmtInf");


                foreach (XmlNode node in nodesBankRef)
                {


                    // MessageBox.Show(node["Ustrd"].InnerText);
                    row.Cells[5].Value = node["Ustrd"].InnerText;

                }
                DBConnect sd = new DBConnect();
              
                row.Cells[6].Value = sd.fetchMail(row.Cells[2].Value.ToString()).ToString();
            }




            foreach (DataGridViewRow rows in m_dataGridViewMailList.Rows)
               // if (row.Index != -1)
                {
                  
                    if (validate_emailaddress.IsMatch(row.Cells[6].Value.ToString()) != true)
                    {

                        row.DefaultCellStyle.BackColor = Color.Red;

                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = Color.White;
                     

                    }
                }

            m_dataGridViewMailList.Rows.Add(row);
            if (row.Index != -1)
            {
                row.Cells[TemplateName.Name].Value = "Template -1";
                tabControl1.SelectedTab = tabControl1.TabPages["tabPage3"];

                row.DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 8);
            }
            m_dataGridViewMailList.AllowUserToAddRows = false;





        }


        private void button3_Click(object sender, EventArgs e)
        {



        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage1"];
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void m_btnDbParametersApply_Click(object sender, EventArgs e)
        {



            config.AppSettings.Settings["Server"].Value = m_txtHost.Text;
            config.AppSettings.Settings["Uid"].Value = m_txtUid.Text;
            config.AppSettings.Settings["Password"].Value = m_txtPassword.Text;
            config.AppSettings.Settings["Database"].Value = m_txtDatabaseName.Text;
            config.Save(ConfigurationSaveMode.Modified);

        }
        bool IsValidEmail(string eMail)
        {
            bool Result = false;

            try
            {
                var eMailValidator = new System.Net.Mail.MailAddress(eMail);

                Result = (eMail.LastIndexOf(".") > eMail.LastIndexOf("@"));
            }
            catch
            {
                Result = false;
            };

            return Result;
        }
        private void m_btnConnect_Click(object sender, EventArgs e)
        {

            // m_dataGridViewMailList.Rows[m_dataGridViewMailList.CurrentRow.Index].Cells[7].Value = "ahmet";

            DBConnect db = new DBConnect();

            if (db.OpenConnection())
            {
                m_btnConnect.Image = WindowsFormsApp2.Properties.Resources.Misc_Web_Database_icon;
                m_lblDbStatus.ForeColor = Color.GreenYellow;
                m_lblDbStatus.Font = new Font("Arial", 14, FontStyle.Bold);
                m_lblDbStatus.Text = "CONNECTED";
                // m_lblDbStatus.Font = new Font(m_lblDbStatus.Font.FontFamily, 16);
                m_btnSettings.Enabled = true;
                m_btnOpenFile.Enabled = true;


                m_txtLog.AppendText("\r\n" + "Connected" + "    " + DateTime.Now.ToString("dd:MM:yyyy HH:mm:ss "));
                m_txtLog.ScrollToCaret();


                {

                };
            }

            else
            {
                m_btnConnect.Image = WindowsFormsApp2.Properties.Resources.Globe_Disconnect_icon;
                m_lblDbStatus.ForeColor = Color.Red;
                m_lblDbStatus.Font = new Font(m_lblDbStatus.Font, FontStyle.Bold);
                m_lblDbStatus.Text = "DISCONNECTED";
                m_btnSettings.Enabled = true;
                m_btnOpenFile.Enabled = false;

                m_txtLog.AppendText("\r\n" + "Disconnected" + "    " + DateTime.Now.ToString("dd:MM:yyyy HH:mm:ss ")); m_txtLog.ScrollToCaret();


            }

        }

        private void m_lblDbStatus_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {
                e.Handled = true;
            }


            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') == -1))
            {
                e.Handled = true;
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            //DBConnect sd = new DBConnect();
            // MessageBox.Show(sd.fetchMail("67").ToString());

            DBConnect testConnect = new DBConnect(m_txtHost.Text, m_txtDatabaseName.Text, m_txtUid.Text, m_txtPassword.Text, m_txtPort.Text, m_txtTimeout.Text);
            if (testConnect.OpenConnection())
            {
                MessageBox.Show("Connection SUCCESFULLY ");
            }
            else
            {
                MessageBox.Show("NOT CONNECTED");
            }

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            m_txtEmail.Enabled = true;

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage3"];


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void m_txtEmail_TextChanged(object sender, EventArgs e)
        {

            m_btnSaveMail.Visible = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message.From = new MailAddress("test@karan.net.tr");
                message.To.Add(new MailAddress("yavuz@karan.net.tr"));
                message.Subject = "Test" + m_txtEmail.Text;
                message.Body = "Content";

                message.SubjectEncoding = System.Text.Encoding.UTF8;

                // set body-message and encoding
                message.Body = "<b>Test Mail</b><br>using <b>HTML" + m_txtHost.Text + "</b>.";
                message.BodyEncoding = System.Text.Encoding.UTF8;
                // text or html
                message.IsBodyHtml = true;

                smtp.Port = 587;
                smtp.Host = "mail.karan.net.tr";
                smtp.EnableSsl = false;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential("test@karan.net.tr", "35Ar6047");
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("err: " + ex.Message);
            }
        }


        private void button5_Click_1(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {

                m_txtLogPath.Text = fbd.SelectedPath;

            }
        }

        private void m_txtLogPath_MouseMove(object sender, MouseEventArgs e)
        {
            toolTip.SetToolTip(m_txtLogPath, m_txtLogPath.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            config.AppSettings.Settings["SMTPHOST"].Value = m_txtSMTPHost.Text;
            config.AppSettings.Settings["SMTPPort"].Value = m_txtSMTPPort.Text;
            config.AppSettings.Settings["SMTPUserName"].Value = m_txtSMTPUserName.Text;
            config.AppSettings.Settings["SMTPPassword"].Value = m_txtSMTPPassword.Text;
            config.AppSettings.Settings["SMTPUseSSL"].Value = m_checkBoxSSL.Checked.ToString();
            config.Save(ConfigurationSaveMode.Modified);
        }

        private void m_btnAddList_Click(object sender, EventArgs e)
        {

            if (validate_emailaddress.IsMatch(m_txtEmail.Text) != true)
            {
                MessageBox.Show("Invalid Email Address!", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                m_txtEmail.Focus();
                return;
            }
            else
            {


            }

        }

        private void m_btnSettings_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage1"];
        }

        private void m_dataGridViewMailList_SelectionChanged(object sender, EventArgs e)
        {
            m_btnSaveMail.Visible = false;
            m_txtEmail.Visible = false;
            m_lblemail.Visible = false;
        }
        int selectedRowindex;
        private void m_dataGridViewMailList_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void m_btnSaveMail_Click(object sender, EventArgs e)
        {

            if (validate_emailaddress.IsMatch(m_txtEmail.Text) != true)
            {
                MessageBox.Show("Invalid Email Address!", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                m_txtEmail.Focus();
                // return;
            }
            else
            {

                m_dataGridViewMailList[6, selectedRowindex].Value = m_txtEmail.Text;


                {
                    DBConnect db = new DBConnect();
                    db.updateMailAddress(row.Cells[2].Value.ToString(), m_txtEmail.Text);
                    m_dataGridViewMailList.Rows[selectedRowindex].DefaultCellStyle.BackColor = Color.White;
                    // m_dataGridViewMailList.Rows[selectedRowindex].DefaultCellStyle.Font= new Font("	Candara Italic", 10, FontStyle.Bold);
                    //   MessageBox.Show(m_dataGridViewMailList.Rows[selectedRowindex].DefaultCellStyle.Font.ToString());
                    m_dataGridViewMailList.Rows[selectedRowindex].DefaultCellStyle.Font = new Font("Candara Italic", 9, FontStyle.Bold);
                }


            }
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {


            for (int i = 0; i < m_dataGridViewMailList.Rows.Count; i++)
            {

                if (validate_emailaddress.IsMatch(m_dataGridViewMailList.Rows[i].Cells[6].Value.ToString()) == true)
                {

                    try
                    {
                        //dbgm@coiver.it

                        MailMessage message = new MailMessage();
                        SmtpClient smtp = new SmtpClient();
                     
                        message.From = new MailAddress(m_txtSMTPUserName.Text);

                        message.To.Add(new MailAddress(m_dataGridViewMailList.Rows[i].Cells[6].Value.ToString()));
                        message.Subject = "Bonifico -> " + row.Cells[1].Value;
                        message.Body = "Content";
                        if (config.AppSettings.Settings["BCC"].Value.ToString() != "")
                        {
                            message.Bcc.Add(config.AppSettings.Settings["BCC"].Value.ToString());
                        }
                        if (config.AppSettings.Settings["CC"].Value.ToString() != "")
                        {
                            message.CC.Add(config.AppSettings.Settings["CC"].Value.ToString());
                        }
                        message.SubjectEncoding = System.Text.Encoding.UTF8;

                        // set body-message and encoding
                        message.Body = "<br><br>Gentile Fornitore  " + row.Cells[1].Value + "," + " <br><br> La presente per informarla che in data odierna la ns. società  ha disposto la seguente disposizione di pagamento : " + row.Cells[1].Value + "<br><br>Vs. C/C : " + row.Cells[4].Value + "<br><br> Importo Euro : " + row.Cells[3].Value + "</b><br><br>Causale pagamento : " + row.Cells[5].Value + "<br><br>Cordiali saluti. " + "<br><br>" + row.Cells[0].Value + "<br><br>--------------------------------------------" + "<br><br>"+ "Amministrazione"+ "<br><br>Gruppo Coiver"+ "<br><br>Per informazioni aggiuntive, chiarire compensazioni, scadenziari si prega scrivere direttamente a"+ "<br>fornitori@coiver.it"+ "<br><br> In alternativa conttatateci telefonicamente tutti i mercoledi mattina al numero Tel +39 02.66.30.18.99";
                        message.BodyEncoding = System.Text.Encoding.UTF8;
                        // text or html
                        message.IsBodyHtml = true;

                        /*
                         
                         config.AppSettings.Settings["SMTPUserName"].Value;
            m_txtSMTPPassword.Text = config.AppSettings.Settings["SMTPPassword"].Value;
            m_checkBoxSSL.Checked = bool.Parse(config.AppSettings.Settings["SMTPUseSSL"].Value);
                         
                         
                         */

                        smtp.Port = Convert.ToInt32(config.AppSettings.Settings["SMTPPort"].Value.ToString());
                        smtp.Host = config.AppSettings.Settings["SMTPHost"].Value;
                        smtp.EnableSsl = bool.Parse(config.AppSettings.Settings["SMTPUseSSL"].Value);
                        smtp.UseDefaultCredentials = false;
                        smtp.Credentials = new NetworkCredential(config.AppSettings.Settings["SMTPUserName"].Value, config.AppSettings.Settings["SMTPPassword"].Value);
                        smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                        smtp.Send(message);
                        m_dataGridViewMailList.Rows[i].DefaultCellStyle.BackColor = Color.GreenYellow;

                        System.IO.File.AppendAllText(m_txtLogPath.Text + "\\" + DateTime.Now.ToString("yyyy_MM_dd") + ".txt", DateTime.Now.ToString("dd/MM/yyyy HH:mm ") + "  Payer: " + m_dataGridViewMailList.Rows[i].Cells[0].Value.ToString() + "  Beneficiary: " + m_dataGridViewMailList.Rows[i].Cells[1].Value.ToString() + "  email : " + m_dataGridViewMailList.Rows[i].Cells[6].Value.ToString() + "  Status: Send " + Environment.NewLine);
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("err: " + ex.Message);
                    }
                }
                else {
                    System.IO.File.AppendAllText(m_txtLogPath.Text + "\\" + DateTime.Now.ToString("yyyy_MM_dd") + ".txt", DateTime.Now.ToString("dd/MM/yyyy HH:mm ") + "  Payer: " + m_dataGridViewMailList.Rows[i].Cells[0].Value.ToString() + "  Beneficiary: " + m_dataGridViewMailList.Rows[i].Cells[1].Value.ToString() + "  email : " + m_dataGridViewMailList.Rows[i].Cells[6].Value.ToString() + "  Status: NOT SEND " + Environment.NewLine);
                }
            }
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void panel11_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void m_dataGridViewMailList_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            selectedRowindex = e.RowIndex;

            if (m_dataGridViewMailList.CurrentCell.ColumnIndex.Equals(6) && e.RowIndex != -1)
            {
                if (m_dataGridViewMailList.CurrentCell != null && m_dataGridViewMailList.CurrentCell.Value != null)
                {
                    m_btnSaveMail.Visible = true;
                    m_txtEmail.Text = m_dataGridViewMailList.Rows[e.RowIndex].Cells[6].Value.ToString();
                    m_txtEmail.Enabled = true;
                    // MessageBox.Show("--------"+m_dataGridViewMailList.Rows[e.RowIndex].Cells[5].Value.ToString());
                    m_lblemail.Visible = true;
                    m_txtEmail.Visible = true;
                }
            }
        }

        private void m_dataGridViewMailList_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {


                int currentMouseOverRow = m_dataGridViewMailList.HitTest(e.X, e.Y).RowIndex;

                if (currentMouseOverRow > -1)
                {
                    ContextMenu m = new ContextMenu();
                    m.MenuItems.Add(new MenuItem("Remove Selected Row", new EventHandler(removeSelected)));
                    m.MenuItems.Add(new MenuItem("Remove Sended List", new EventHandler(removeSendedList)));
                    m.MenuItems.Add(new MenuItem("Remove UnSended List", new EventHandler(removeUnsendedList)));
                    m.MenuItems.Add(new MenuItem("Remove  ALL ", new EventHandler(removeAll)));
                   
                    m.Show(m_dataGridViewMailList, new Point(e.X, e.Y));
                    // m.MenuItems.Add(new MenuItem(string.Format("Do something to row {0}", currentMouseOverRow.ToString())));
                }



            }

        }
        private void removeSelected(object sender, EventArgs e)
        {
           // Int32 rowToDelete = m_dataGridViewMailList.Rows.GetFirstRow(DataGridViewElementStates.Selected);
     
            //m_dataGridViewMailList.Rows.RemoveAt(rowToDelete);
           // m_dataGridViewMailList.ClearSelection();
        }
        private void removeAll(object sender, EventArgs e) { 
        m_dataGridViewMailList.Rows.Clear();
        }
        private void removeUnsendedList(object sender, EventArgs e)
        {

            foreach (DataGridViewRow rows in m_dataGridViewMailList.Rows)

                for (int i = 0; i < m_dataGridViewMailList.Rows.Count; i++)
                {

                    if (m_dataGridViewMailList.Rows[i].DefaultCellStyle.BackColor == Color.Red)
                    {


                        m_dataGridViewMailList.Rows.RemoveAt(m_dataGridViewMailList.Rows[i].Index);




                        // MessageBox.Show((m_dataGridViewMailList.Rows[i].Index).ToString());

                    }

                }



            foreach (DataGridViewRow rows in m_dataGridViewMailList.Rows)

                for (int i = 0; i < m_dataGridViewMailList.Rows.Count; i++)
                {

                    if (m_dataGridViewMailList.Rows[i].DefaultCellStyle.BackColor == Color.Red)
                    {


                        m_dataGridViewMailList.Rows.RemoveAt(m_dataGridViewMailList.Rows[i].Index);




                        // MessageBox.Show((m_dataGridViewMailList.Rows[i].Index).ToString());

                    }

                }
        }
        private void removeSendedList(object sender, EventArgs e)
        {
     
        
            foreach (DataGridViewRow rows in m_dataGridViewMailList.Rows)
      
            for (int i = 0; i < m_dataGridViewMailList.Rows.Count; i++)
                {
               
                        if (m_dataGridViewMailList.Rows[i].DefaultCellStyle.BackColor == Color.GreenYellow)
                    {
                       

                        m_dataGridViewMailList.Rows.RemoveAt(m_dataGridViewMailList.Rows[i].Index);
                       



                        // MessageBox.Show((m_dataGridViewMailList.Rows[i].Index).ToString());

                    }
                    
                }


         
            foreach (DataGridViewRow rows in m_dataGridViewMailList.Rows)

                for (int i = 0; i < m_dataGridViewMailList.Rows.Count; i++)
                {

                    if (m_dataGridViewMailList.Rows[i].DefaultCellStyle.BackColor == Color.GreenYellow)
                    {


                        m_dataGridViewMailList.Rows.RemoveAt(m_dataGridViewMailList.Rows[i].Index);




                        // MessageBox.Show((m_dataGridViewMailList.Rows[i].Index).ToString());

                    }

                }

        }



        private void m_dataGridViewMailList_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            /*
            m_dataGridViewMailList.BeginEdit(false);

            bool validClick = (e.RowIndex != -1 && e.ColumnIndex != -1); //Make sure the clicked row/column is valid.
            var datagridview = sender as DataGridView;

            // Check to make sure the cell clicked is the cell containing the combobox 
            if (datagridview.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && validClick)
            {
                m_dataGridViewMailList.BeginEdit(true);
                ComboBox com = (ComboBox)this.m_dataGridViewMailList.EditingControl;
                com.DroppedDown = true;
            }


            */


            //com.BackColor=Color.White;

            if (e.RowIndex != -1)
            {
                if (m_dataGridViewMailList.Rows[e.RowIndex].Cells[1].Value != null)
                {


                    if (m_dataGridViewMailList.CurrentCell.ColumnIndex.Equals(7) && e.RowIndex != -1)
                    {
                        if (m_dataGridViewMailList.Rows[0].Cells[TemplateName.Name].Value != null && m_dataGridViewMailList.CurrentCell.Value != null)
                        {
                            return;


                        }

                        if (e.RowIndex < 0)
                        {
                            return;
                        }

                        m_dataGridViewMailList.BeginEdit(true);
                        ComboBox com = (ComboBox)this.m_dataGridViewMailList.EditingControl;
                        com.DroppedDown = true;
                        com.BackColor = Color.White;

                    }
                }
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {




        }

        private void m_dataGridViewMailList_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is ComboBox)
            {
                ComboBox comboBox = e.Control as ComboBox;
                comboBox.SelectedIndexChanged -= LastColumnComboSelectionChanged;
                comboBox.SelectedIndexChanged += LastColumnComboSelectionChanged;
            }

        }
        private void LastColumnComboSelectionChanged(object sender, EventArgs e)
        {
            var currentcell = m_dataGridViewMailList.CurrentCellAddress;
            var sendingCB = sender as DataGridViewComboBoxEditingControl;
            DataGridViewTextBoxCell cel = (DataGridViewTextBoxCell)m_dataGridViewMailList.Rows[currentcell.Y].Cells[0];
            cel.Value = sendingCB.EditingControlFormattedValue.ToString();
            // string Code = ((DataRowView)m_dataGridViewMailList.SelectedItem).Row["Code"].ToString();
            MessageBox.Show(cel.Value.ToString());
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            config.AppSettings.Settings["UseLogFile"].Value = checkBox2.Checked.ToString();
            config.AppSettings.Settings["LogFilePath"].Value = m_txtLogPath.Text;

            config.Save(ConfigurationSaveMode.Modified);

        }

        private void m_txtPassword_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabControl1_KeyDown(object sender, KeyEventArgs e)
        {
        
        }

        private void m_dataGridViewMailList_MouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                this.m_dataGridViewMailList.Rows[e.RowIndex].Selected = true;
                MessageBox.Show(e.RowIndex.ToString());
           

            }
        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            config.AppSettings.Settings["CC"].Value = m_txtcc.Text;
            config.AppSettings.Settings["BCC"].Value = m_txtbcc.Text;
           
            config.Save(ConfigurationSaveMode.Modified);
        }
    }
}

