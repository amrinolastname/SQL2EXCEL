using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SQL2XL
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void buttonSubmit_Click(object sender, EventArgs e)
        {
            this.Close();
            sql.connectionString = "server=" + textBoxServer.Text + ";uid=" + textBoxUser.Text + ";pwd=" + textBoxPassword.Text + ";database=" + textBoxDatabase.Text;
            Properties.Settings.Default.server = textBoxServer.Text;
            Properties.Settings.Default.database = textBoxDatabase.Text;
            Properties.Settings.Default.user = textBoxUser.Text;
            Properties.Settings.Default.password = textBoxPassword.Text;
            Properties.Settings.Default.Save();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            var settings = Properties.Settings.Default;
            textBoxServer.Text = settings.server;
            textBoxDatabase.Text = settings.database;
            textBoxUser.Text = settings.user;
            textBoxPassword.Text = settings.password;
        }
    }
}
