using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Peppermint_Outlook_AddIn
{
    public partial class frmAbout : Form
    {
        public frmAbout()
        {
            InitializeComponent();
        }

        private void frmAbout_Load(object sender, EventArgs e)
        {
            labelVersion.Text = "Version : " + Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        private void label2_Click(object sender, EventArgs e)
        {
            Process.Start("http://Peppermint.com");
        }
    }
}
