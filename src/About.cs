using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace swsp
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();
            this.versionLabel.Text = String.Format(this.versionLabel.Text, Assembly.GetExecutingAssembly().GetName().Version.ToString().Substring(0, 5));
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("http://traal.eu/swsp");
        }
    }
}
