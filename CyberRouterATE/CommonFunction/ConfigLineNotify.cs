///---------------------------------------------------------------------------------------
///  This code was created by CyberTan James Chu.
///  File           : ConfigSerialPort.cs
///  Update         : 2014-07-22
///  Version        : 1.0.140722
///  Description    : 
///  Modified       : 2014-07-22 Initial version
///---------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Xml;
using System.Diagnostics;

namespace CyberRouterATE
{
    public partial class ConfigLineNotify : Form
    {
        /// =======================
        /// *******Event*******
        /// =======================

        public ConfigLineNotify()
        {
            InitializeComponent();
        }

        private void ConfigLineNotify_Load(object sender, EventArgs e)
        {
            cbLineTestItem.SelectedIndex = 0;

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (btnSave.Enabled)
            {

                try
                {
                    WriteXmlLineNotify(System.Windows.Forms.Application.StartupPath + @"\config\Line_Notify_" + cbLineTestItem.SelectedItem.ToString() + @".xml");
                }
                catch { }
            }

            this.Close();
        }

        private void cbLineTestItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            try
            {
                ReadXmlLineNotify(System.Windows.Forms.Application.StartupPath + @"\config\Line_Notify_" + cbLineTestItem.SelectedItem.ToString() + @".xml", cbLineTestItem.SelectedItem.ToString());
            }
            catch
            {
                for (int i = 0; i <= (clboxConfigLineNotifyGroup.Items.Count - 1); i++)
                {
                    clboxConfigLineNotifyGroup.SetItemChecked(i, false);
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string sTestItem = cbLineTestItem.SelectedItem.ToString();

            try
            {
                WriteXmlLineNotify(System.Windows.Forms.Application.StartupPath + @"\config\Line_Notify_" + cbLineTestItem.SelectedItem.ToString() + @".xml");
            }
            catch { }

            btnSave.Enabled = false;
        }

        private void clboxConfigLineNotifyGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSave.Enabled = true;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string sTestItem = "Line_Notify_" + cbLineTestItem.SelectedItem.ToString() + "_Export";

            // Displays a SaveFileDialog so the user can save the XML assigned to Save config.
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\config\";
            saveFileDialog1.FileName = sTestItem;
            saveFileDialog1.DefaultExt = ".xml";
            saveFileDialog1.Filter = "XML file|*.xml";
            saveFileDialog1.Title = "Save an xml file";

            // If the file name is not an empty string open it for saving.
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName != "")
            {
                WriteXmlLineNotify(saveFileDialog1.FileName);
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            string filename = "Line_Notify_" + cbLineTestItem.SelectedItem.ToString() + "_Export";


            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = filename;
            openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\config\";
            // Set filter for file extension and default file extension
            openFileDialog1.Filter = "XML file|*.xml";

            // If the file name is not an empty string open it for opening.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName != "")
            {
                ReadXmlLineNotify(openFileDialog1.FileName, cbLineTestItem.SelectedItem.ToString());
            }
        }



        /// =======================
        /// *******Function*******
        /// =======================

        private void WriteXmlLineNotify(string filename)
        {
            XmlWriterSettings setting = new XmlWriterSettings();
            setting.Indent = true; //指定縮排

            XmlWriter writer = XmlWriter.Create(filename, setting);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by Line Notify.");
            writer.WriteStartElement("CybertanATE");
            writer.WriteAttributeString("Item", cbLineTestItem.SelectedItem.ToString());


            /*Write Groups*/
            writer.WriteStartElement("Groups");

            for (int i = 0; i <= (clboxConfigLineNotifyGroup.Items.Count - 1); i++)
            {
                if (clboxConfigLineNotifyGroup.GetItemChecked(i))
                {
                    writer.WriteElementString("ATE_Notify_" + (i + 1).ToString(), "true");
                }
                else
                {
                    writer.WriteElementString("ATE_Notify_" + (i + 1).ToString(), "false");
                }
            }

            writer.WriteEndElement();
            /* Writer */


            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
            btnSave.Enabled = false;
        }

        private bool ReadXmlLineNotify(string filename, string nodeName)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(filename);

            XmlNode node = doc.SelectSingleNode("CybertanATE");
            if (node == null)
            {
                return false;
            }

            XmlElement element = (XmlElement)node;
            string strID = element.GetAttribute("Item");
            Debug.Write(strID);
            if (strID.CompareTo(nodeName) != 0)
            {
                MessageBox.Show("This XML file is incorrect", "Error");
                return false;
            }

            XmlNode nodeAttenuator = doc.SelectSingleNode("/CybertanATE/Groups");

            try
            {
                string ATE_Notify_1 = nodeAttenuator.SelectSingleNode("ATE_Notify_1").InnerText;
                string ATE_Notify_2 = nodeAttenuator.SelectSingleNode("ATE_Notify_2").InnerText;
                string ATE_Notify_3 = nodeAttenuator.SelectSingleNode("ATE_Notify_3").InnerText;
                string ATE_Notify_4 = nodeAttenuator.SelectSingleNode("ATE_Notify_4").InnerText;
                string ATE_Notify_5 = nodeAttenuator.SelectSingleNode("ATE_Notify_5").InnerText;
                string ATE_Notify_6 = nodeAttenuator.SelectSingleNode("ATE_Notify_6").InnerText;

                Debug.WriteLine("ATE_Notify_1: " + ATE_Notify_1);
                Debug.WriteLine("ATE_Notify_2: " + ATE_Notify_2);
                Debug.WriteLine("ATE_Notify_3: " + ATE_Notify_3);
                Debug.WriteLine("ATE_Notify_4: " + ATE_Notify_4);
                Debug.WriteLine("ATE_Notify_5: " + ATE_Notify_5);
                Debug.WriteLine("ATE_Notify_6: " + ATE_Notify_6);

                if (ATE_Notify_1 == "true")
                    clboxConfigLineNotifyGroup.SetItemChecked(0, true);
                else
                    clboxConfigLineNotifyGroup.SetItemChecked(0, false);

                if (ATE_Notify_2 == "true")
                    clboxConfigLineNotifyGroup.SetItemChecked(1, true);
                else
                    clboxConfigLineNotifyGroup.SetItemChecked(1, false);

                if (ATE_Notify_3 == "true")
                    clboxConfigLineNotifyGroup.SetItemChecked(2, true);
                else
                    clboxConfigLineNotifyGroup.SetItemChecked(2, false);

                if (ATE_Notify_4 == "true")
                    clboxConfigLineNotifyGroup.SetItemChecked(3, true);
                else
                    clboxConfigLineNotifyGroup.SetItemChecked(3, false);

                if (ATE_Notify_5 == "true")
                    clboxConfigLineNotifyGroup.SetItemChecked(4, true);
                else
                    clboxConfigLineNotifyGroup.SetItemChecked(4, false);

                if (ATE_Notify_6 == "true")
                    clboxConfigLineNotifyGroup.SetItemChecked(5, true);
                else
                    clboxConfigLineNotifyGroup.SetItemChecked(5, false);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }


            return true;
        }









    }
}
