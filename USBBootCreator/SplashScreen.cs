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

namespace USBBootCreator
{
    public partial class SplashScreen : Form
    {
        public SplashScreen()
        {
            InitializeComponent();
            this.progressBar1.Maximum = 100;
        }

        private void SplashScreen_Load(object sender, EventArgs e)
        {
            var productInfo = FileVersionInfo.GetVersionInfo(Assembly.GetEntryAssembly().Location);
            label1.Text = Application.ProductVersion + " " + productInfo.LegalCopyright;
        }
    }
}
