using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace CertEnroll
{
    public partial class Configuration : Form
    {
        public Configuration()
        {
            InitializeComponent();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Yes or no", "Save and exit?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes)
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("Config.xml");
                doc.SelectSingleNode("/appSettings/configuration/OrganizationName").InnerText = textBox1.Text;
                doc.SelectSingleNode("/appSettings/configuration/DomainName").InnerText = textBox2.Text;
                doc.SelectSingleNode("/appSettings/configuration/CAName").InnerText = textBox3.Text;
                doc.SelectSingleNode("/appSettings/configuration/Template").InnerText = textBox4.Text;
                doc.SelectSingleNode("/appSettings/configuration/ExportPass").InnerText = CryptUtils.EncryptString(textBox5.Text, CryptUtils.configPassword);
                //richTextBox1.Lines = doc.SelectSingleNode("/appSettings/configuration/OU").InnerText.Split(',');
                //doc.SelectSingleNode("/appSettings/configuration/OU").InnerText = String.Join(",", richTextBox1.Lines);
                string OUtext = "";
                for (int i=0;i<dataGridView1.Rows.Count-1;i++)
                {
                    OUtext = OUtext + dataGridView1.Rows[i].Cells["OU"].Value + ",";
                }
                doc.SelectSingleNode("/appSettings/configuration/OU").InnerText = OUtext.Remove(OUtext.Length - 1);
                doc.SelectSingleNode("/appSettings/configuration/wwwEnable").InnerText = Convert.ToString(checkBox1.Checked);
                doc.SelectSingleNode("/appSettings/configuration/wwwPort").InnerText = numericUpDown1.Text;
                doc.SelectSingleNode("/appSettings/configuration/wwwPath").InnerText = label1.Text;
                doc.SelectSingleNode("/appSettings/configuration/wwwIP").InnerText = comboBox1.Text;
                doc.Save("Config.xml");
                this.Close();
            }
        }

        private void Configuration_Load(object sender, EventArgs e)
        {
            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

            columnHeaderStyle.BackColor = Color.Beige;
            columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
            dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            XmlDocument doc = new XmlDocument();
            doc.Load("Config.xml");
            textBox1.Text = doc.SelectSingleNode("/appSettings/configuration/OrganizationName").InnerText;
            textBox2.Text = doc.SelectSingleNode("/appSettings/configuration/DomainName").InnerText;
            textBox3.Text = doc.SelectSingleNode("/appSettings/configuration/CAName").InnerText;
            textBox4.Text = doc.SelectSingleNode("/appSettings/configuration/Template").InnerText;
            textBox5.Text = CryptUtils.DecryptString(doc.SelectSingleNode("/appSettings/configuration/ExportPass").InnerText, CryptUtils.configPassword);
            dataGridView1.Columns.Add("OU","OU");
            foreach (var el in doc.SelectSingleNode("/appSettings/configuration/OU").InnerText.Split(','))
            {
                dataGridView1.Rows.Add(el.ToString());
            }
            checkBox1.Checked = Convert.ToBoolean(doc.SelectSingleNode("/appSettings/configuration/wwwEnable").InnerText);

            if (checkBox1.Checked)
            {
                comboBox1.Enabled = true;
                numericUpDown1.Enabled = true;
                label1.Enabled = true;
                label7.Enabled = true;
                label8.Enabled = true;
                button3.Enabled = true;
                //comboBox1.Items.Add(doc.SelectSingleNode("/appSettings/configuration/wwwIP").InnerText);
                foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
                {
                    foreach (UnicastIPAddressInformation ip in nic.GetIPProperties().UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        {
                            comboBox1.Items.Add(ip.Address.ToString());
                        }
                    }
                }
                //comboBox1.SelectedIndex = 0;
                comboBox1.SelectedItem = doc.SelectSingleNode("/appSettings/configuration/wwwIP").InnerText;
                numericUpDown1.Text = doc.SelectSingleNode("/appSettings/configuration/wwwPort").InnerText;
                label1.Text = doc.SelectSingleNode("/appSettings/configuration/wwwPath").InnerText;
            }
            else
            {
                comboBox1.Enabled = false;
                numericUpDown1.Enabled = false;
                label1.Enabled = false;
                label7.Enabled = false;
                label8.Enabled = false;
                button3.Enabled = false;
                //comboBox1.Items.Add(doc.SelectSingleNode("/appSettings/configuration/wwwIP").InnerText);
                foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
                {
                    foreach (UnicastIPAddressInformation ip in nic.GetIPProperties().UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        {
                            comboBox1.Items.Add(ip.Address.ToString());
                        }
                    }
                }
                //comboBox1.SelectedIndex = 0;
                comboBox1.SelectedItem = doc.SelectSingleNode("/appSettings/configuration/wwwIP").InnerText;
                numericUpDown1.Text = doc.SelectSingleNode("/appSettings/configuration/wwwPort").InnerText;
                label1.Text = doc.SelectSingleNode("/appSettings/configuration/wwwPath").InnerText;
            }
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                comboBox1.Enabled = true;
                numericUpDown1.Enabled = true;
                label1.Enabled = true;
                label7.Enabled = true;
                label8.Enabled = true;
                button3.Enabled = true;
             }
            else
            {
                comboBox1.Enabled = false;
                numericUpDown1.Enabled = false;
                label1.Enabled = false;
                label7.Enabled = false;
                label8.Enabled = false;
                button3.Enabled = false;
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                label1.Text = folderBrowserDialog1.SelectedPath;
            }
        }
    }
}
