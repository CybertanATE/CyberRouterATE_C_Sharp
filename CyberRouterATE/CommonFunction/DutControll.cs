using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using RouterControlClass;

namespace CyberRouterATE
{
    public partial class DutControll : Form
    {
        public string DutConfigurationType = string.Empty;
        public string DutConfigurationlFile = string.Empty;
        RouterGuiControl guiControl = null;

        public DutControll()
        {
            InitializeComponent();            
        }
        
        private void btnOK_Click(object sender, EventArgs e)
        {
            DutConfigurationType = cboxDutConfigurationType.Text;
            DutConfigurationlFile = txtDutConfigurationFile.Text;

            guiControl = new RouterGuiControl(cboxDutConfigurationType.Text, txtDutConfigurationFile.Text);

            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            DutConfigurationType = string.Empty;
            DutConfigurationlFile = string.Empty;

            this.Close();
        }

        private void txtDutConfigurationFile_Click(object sender, EventArgs e)
        {
            string path = System.Windows.Forms.Application.StartupPath;
            path = path + "\\DutConfiguration";
            if (!File.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = "DeviceConfiguration.xml";
            openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\config\";
            // Set filter for file extension and default file extension
            openFileDialog1.Filter = "XML file|*.xml";

            // If the file name is not an empty string open it for opening.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName != "")
            {
                txtDutConfigurationFile.Text = openFileDialog1.FileName;
            }
        }
    }
}
