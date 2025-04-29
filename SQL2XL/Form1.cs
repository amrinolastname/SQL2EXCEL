using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SQL2XL
{
    public partial class Form1 : Form
    {
        public Form formConnection = new Form2();
        public Form1()
        {
            InitializeComponent();

            var settings = Properties.Settings.Default;

            if (settings.server == "" || settings.database == "" || settings.user == "") sql.connectionString = "";
            else sql.connectionString = "server=" + settings.server + ";uid=" + settings.user + ";pwd=" + settings.password + ";database=" + settings.database;


            formConnection.StartPosition = FormStartPosition.Manual;

            queryTextBox.Select();

        }

        private void buttonExecute_Click(object sender, EventArgs e)
        {
            executeSQL();
        }
        private void executeSQL()
        {
            if (sql.connectionString == "") setDBConnection();
            if (sql.connectionString == "") return;
            setStatus("Reading from database");
            int rowCount = sql.readToGrid(queryTextBox.Text, dataGridView1, 1, 0, null);
            setStatus("Retrieved " + rowCount.ToString() + " rows");
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            exportToExel();
        }
        private void exportToExel()
        {
            setStatus("Exporting to Excel",false);
            excel.exportToExel(dataGridView1, false);
        }
        private void buttonExecuteAndExport_Click(object sender, EventArgs e)
        {
            executeSQL();
            exportToExel();
        }

        public void setStatus(string message,bool deleteOldMessage=true)
        {
            if (!deleteOldMessage) message = statusLabel.Text + "\n" + message;
            statusLabel.Text = message;
            statusLabel.Invalidate();
            statusLabel.Update();
        }

        private void connectionButton_Click(object sender, EventArgs e)
        {
            setDBConnection();
        }

        private void setDBConnection()
        {
            formConnection.Top = this.Top + connectionButton.Top;
            formConnection.Left = this.Left + connectionButton.Left;
            formConnection.ShowDialog();
        }
    }
}
