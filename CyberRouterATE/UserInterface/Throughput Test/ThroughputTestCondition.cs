///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : ThroughputTestCondition.cs
///  Update         : 2017-03-30
///  Description    : All Channel Throughput Test Condtion 
///  Modified       : 2017-03-30 Initial version  
///
///  Comments       : 
///---------------------------------------------------------------------------------------
///

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Diagnostics;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        /* Declare global variable */        
        //string[] Mode_2_4G = {"20M", "40M"};
        //string[] Mode_5G = { "11N_20M", "11N_40M", "11AC_20M", "11AC_40M", "11AC_80M" };


        private void InitThroughputTestCondition()
        {
            rbtnThroughputTestCondition24G.Checked = true;
            SetupDataGridViewThroughputTest();           
        }
              

        private void rbtnThroughputTestCondition24G_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnThroughputTestCondition24G.Checked == true)
            {                
                cboxThroughputTestConditionWirelessMode.Items.Clear();

                int i;
                for (i = 0; i < Mode_2_4G.Length; i++)
                {
                    cboxThroughputTestConditionWirelessMode.Items.Add(Mode_2_4G[i]);
                }

                cboxThroughputTestConditionWirelessMode.SelectedIndex = 0;               
            }
            else
            {
                cboxThroughputTestConditionWirelessMode.Items.Clear();

                int i;
                for (i = 0; i < Mode_5G.Length; i++)
                {
                    cboxThroughputTestConditionWirelessMode.Items.Add(Mode_5G[i]);
                }

                cboxThroughputTestConditionWirelessMode.SelectedIndex = 0; 
            }
        }

        public void SetupDataGridViewThroughputTest()
        {
            dgvThroughputTestConditionData.ColumnCount = 12;
            dgvThroughputTestConditionData.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvThroughputTestConditionData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvThroughputTestConditionData.Name = "Test Condition Setting";
            //dgvThroughputTestConditionData.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgvThroughputTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        
            dgvThroughputTestConditionData.Columns[0].Name = "Band";
            dgvThroughputTestConditionData.Columns[1].Name = "Mode";
            dgvThroughputTestConditionData.Columns[2].Name = "SSID Config";
            dgvThroughputTestConditionData.Columns[3].Name = "SSID_Text";
            dgvThroughputTestConditionData.Columns[4].Name = "Channel Config";
            dgvThroughputTestConditionData.Columns[5].Name = "Channel";            
            dgvThroughputTestConditionData.Columns[6].Name = "Security Config";       
            dgvThroughputTestConditionData.Columns[7].Name = "Security Mode";            
            dgvThroughputTestConditionData.Columns[8].Name = "Passphrase";
            dgvThroughputTestConditionData.Columns[9].Name = "TX tst File";
            dgvThroughputTestConditionData.Columns[10].Name = "RX tst File";
            dgvThroughputTestConditionData.Columns[11].Name = "BI-Dir tst File";
            
            /*
            dgvThroughputTestConditionData.Columns[0].Width = 50;
            dgvThroughputTestConditionData.Columns[1].Width = 50;
            dgvThroughputTestConditionData.Columns[2].Width = 50;
            dgvThroughputTestConditionData.Columns[3].Width = 20;
            dgvThroughputTestConditionData.Columns[4].Width = 50;
            dgvThroughputTestConditionData.Columns[5].Width = 20; 
            dgvThroughputTestConditionData.Columns[6].Width = 50;
            dgvThroughputTestConditionData.Columns[7].Width = 120;
            dgvThroughputTestConditionData.Columns[8].Width = 20;
            dgvThroughputTestConditionData.Columns[9].Width = 120;
            dgvThroughputTestConditionData.Columns[10].Width = 120;
            dgvThroughputTestConditionData.Columns[11].Width = 120;
            dgvThroughputTestConditionData.Columns[12].Width = 120;
            dgvThroughputTestConditionData.Columns[13].Width = 120;
            */
            
            dgvThroughputTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvThroughputTestConditionData.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvThroughputTestConditionData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dgvThroughputTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvThroughputTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;


            /*
            string[] row = { (dgvThroughputTestConditionData.NewRowIndex + 1).ToString(), numericUpDown1.Value.ToString(), numericUpDown2.Value.ToString(), numericUpDown3.Value.ToString(), textBox1.Text, textBox2.Text };
            dgvThroughputTestConditionData.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvThroughputTestConditionData.AutoResizeColumns();

            //dgvThroughputTestConditionData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            //dgvThroughputTestConditionData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgvThroughputTestConditionData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
            dgvThroughputTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvThroughputTestConditionData.Rows.Add(row);
            dgvThroughputTestConditionData.FirstDisplayedScrollingRowIndex = dgvThroughputTestConditionData.Rows.Count - 1;
            */        
        }

        private void btnThroughputTestConditionAddSetting_Click(object sender, EventArgs e)
        {
            /* Check all data has value */
            if (txtThroughputTestConditionSSIDText.Text == "")
            {
                MessageBox.Show("SSID Can't be Empty!!", "Error");
                return;
            }

            /* Check at least one tst file has selected */
            if (txtThroughputTestConditionTxTst.Text == "" && txtThroughputTestConditionRxTst.Text == "" && txtThroughputTestConditionBiTst.Text == "")
            {
                MessageBox.Show("Select at least one TST File!!", "Error");
                return;
            }

            //if (nudThroughputTestConditionWirelessChannelStart.Value > nudThroughputTestConditionWirelessChannelStop.Value)
            //{
            //    MessageBox.Show("Channel Stop Value must equal or greater than Start Value!!", "Error");
            //    return;
            //}

            /* Check if Delete button exist in the datagridview */
            if (hasDeleteButton)
            {
                hasDeleteButton = false;
                dgvThroughputTestConditionData.Columns.RemoveAt(0);
                /* Reset the name of btnThroughputTestConditionEditSetting as Edit words. */
                btnThroughputTestConditionEditSetting.Text = "Edit";
            }
           
            /* Add data to datagridview */
            string[] row = new string[] {
                rbtnThroughputTestCondition24G.Checked ? "2.4G" : "5G", //Band
                cboxThroughputTestConditionWirelessMode.Text,  //Mode
                chkThroughputTestConditionWirelessSsid.Checked? "Y" : "N", //SSID_Config
                txtThroughputTestConditionSSIDText.Text, //SSID_Text
                chkThroughputTestConditionWirelessChannel.Checked? "Y" : "N", //Channel_Config
                nudThroughputTestConditionWirelessChannelStart.Value.ToString(), //Channel_Start                
                chkThroughputTestConditionWirelessSecurity.Checked? "Y" : "N",//Security_Config
                cboxThroughputTestConditionSecurity.Text, //Security_Mode            
                txtThroughputTestConditionPassphrase.Text, //Key 
                txtThroughputTestConditionTxTst.Text, //TX tst file
                txtThroughputTestConditionRxTst.Text, //RX tst File
                txtThroughputTestConditionBiTst.Text}; //Bi-Direction tst File*/};  

            dgvThroughputTestConditionData.Rows.Add(row);
            dgvThroughputTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvThroughputTestConditionData.AutoResizeColumns();
            dgvThroughputTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;  
        }

        private void btnThroughputTestConditionEditSetting_Click(object sender, EventArgs e)
        {
            /* James: To prevent one cell is to be inserted more time when Edit button clicked once again. */
            if (hasDeleteButton == false)
            {
                hasDeleteButton = true;
                btnThroughputTestConditionEditSetting.Text = "Cancel";
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.Width = 66;
                btn.HeaderText = "Action";
                btn.Text = "Delete";
                btn.Name = "Action";
                btn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                btn.UseColumnTextForButtonValue = true;

                dgvThroughputTestConditionData.Columns.Insert(0, btn);
            }
            else
            {
                btnThroughputTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvThroughputTestConditionData.Columns.Remove("Action");
            }
        }

        private void btnThroughputTestConditionSaveSetting_Click(object sender, EventArgs e)
        {
            if (dgvThroughputTestConditionData.RowCount <= 1)
            {
                MessageBox.Show("Add Condition First!!", "Warning");
                return;
            }

            string filename = "ThroughputTestCondition";

            // Displays a SaveFileDialog so the user can save the XML assigned to Save config.
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";
            saveFileDialog1.FileName = filename;
            saveFileDialog1.DefaultExt = ".xml";
            saveFileDialog1.Filter = "XML file|*.xml";
            saveFileDialog1.Title = "Save an xml file";

            // If the file name is not an empty string open it for saving.
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName != "")
            {
                writeXmlThroughputTestCondition(saveFileDialog1.FileName);
            }
        }

        private void btnThroughputTestConditionLoadSetting_Click(object sender, EventArgs e)
        {
            if (btnThroughputTestConditionEditSetting.Text.ToLower() == "cancel")
            {
                btnThroughputTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvThroughputTestConditionData.Columns.Remove("Action");
            }

            if (dgvThroughputTestConditionData.RowCount > 1)
            {
                if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                    return;
                else
                {
                    DataTable dt = (DataTable)dgvThroughputTestConditionData.DataSource;
                    dgvThroughputTestConditionData.Rows.Clear();
                    dgvThroughputTestConditionData.DataSource = dt;
                }
            }

            string filename = string.Empty;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = filename;
            openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";
            // Set filter for file extension and default file extension
            openFileDialog1.Filter = "XML file|*.xml";

            // If the file name is not an empty string open it for opening.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName != "")
            {
                readXmlThroughputTestCondition(openFileDialog1.FileName);
            }

        }
        
        private void dgvThroughputTestConditionData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            /* James modified: Adding a condition to prevent exception error when remove null cell. */
            if (dgvThroughputTestConditionData.Columns[e.ColumnIndex].Name == "Action" && hasDeleteButton == true && dgvThroughputTestConditionData.Rows[e.RowIndex].Cells[0].Value != null)
            {
                dgvThroughputTestConditionData.Rows.RemoveAt(e.RowIndex);
                if (e.ColumnIndex == 0 && (String)dgvThroughputTestConditionData.Rows[0].Cells[0].Value == String.Empty)
                    dgvThroughputTestConditionData.Columns.RemoveAt(0);
            }
        }

        private void btnThroughputTestConditionTxTst_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.DefaultExt = "tst";
            openFileDialog1.Title = "Choose Test tst file ";
            openFileDialog1.Filter = "tst files(*.txt)|*.tst|All files(*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "Tx.tst";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtThroughputTestConditionTxTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open TX TST file: " + ex.Message);
                }
            }
        }

        private void btnThroughputTestConditionRxTst_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.DefaultExt = "tst";
            openFileDialog1.Title = "Choose Test tst file ";
            openFileDialog1.Filter = "tst files(*.txt)|*.tst|All files(*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "Rx.tst";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtThroughputTestConditionRxTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open RX TST file: " + ex.Message);
                }
            }
        }

        private void btnThroughputTestConditionBiTst_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.DefaultExt = "tst";
            openFileDialog1.Title = "Choose Test tst file ";
            openFileDialog1.Filter = "tst files(*.txt)|*.tst|All files(*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "TX-RX.tst";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtThroughputTestConditionBiTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open Bi-Direction TST file: " + ex.Message);
                }
            }
        }

        public void writeXmlThroughputTestCondition(string FileName)
        {
            // ToDo: Needs to verify which test item is to be selected..

            string[,] rowdata = new string[dgvThroughputTestConditionData.RowCount - 1, 7];

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;

            XmlWriter writer = XmlWriter.Create(FileName, settings);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by the program.");
            writer.WriteStartElement("CyberRouterATETC");
            writer.WriteAttributeString("Item", "Throughput Test Condition");

            ///
            /// Write Function Test settings
            /// 
            writer.WriteStartElement("TestCondition");
            for (int i = 0; i < dgvThroughputTestConditionData.RowCount - 1; i++)
            {
                writer.WriteStartElement("Condition_" + (i + 1).ToString());

                writer.WriteElementString("Band",           dgvThroughputTestConditionData.Rows[i].Cells[0].Value.ToString());
                writer.WriteElementString("Mode",           dgvThroughputTestConditionData.Rows[i].Cells[1].Value.ToString());
                writer.WriteElementString("SSID_Config",    dgvThroughputTestConditionData.Rows[i].Cells[2].Value.ToString());
                writer.WriteElementString("SSID_Text",      dgvThroughputTestConditionData.Rows[i].Cells[3].Value.ToString());
                writer.WriteElementString("Channel_Config", dgvThroughputTestConditionData.Rows[i].Cells[4].Value.ToString());
                writer.WriteElementString("Channel",        dgvThroughputTestConditionData.Rows[i].Cells[5].Value.ToString());                
                writer.WriteElementString("Security_Config",dgvThroughputTestConditionData.Rows[i].Cells[6].Value.ToString());
                writer.WriteElementString("Security_Mode",  dgvThroughputTestConditionData.Rows[i].Cells[7].Value.ToString());                
                writer.WriteElementString("Passphrase",     dgvThroughputTestConditionData.Rows[i].Cells[8].Value.ToString());
                writer.WriteElementString("TX_tst_File",    dgvThroughputTestConditionData.Rows[i].Cells[9].Value.ToString());
                writer.WriteElementString("RX_tst_File",    dgvThroughputTestConditionData.Rows[i].Cells[10].Value.ToString());
                writer.WriteElementString("BiDir_tst_File", dgvThroughputTestConditionData.Rows[i].Cells[11].Value.ToString());

                writer.WriteEndElement();
            }

            writer.WriteStartElement("Condition_Number");
            writer.WriteElementString("Number", dgvThroughputTestConditionData.RowCount.ToString());
            writer.WriteEndElement();

            // End of write Function Test
            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public bool readXmlThroughputTestCondition(string FileName)
        {
            int number = 0;

            dgvThroughputTestConditionData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            XmlDocument doc = new XmlDocument();
            doc.Load(FileName);

            XmlNode nodeXml = doc.SelectSingleNode("CyberRouterATETC");
            if (nodeXml == null)
                return false;
            XmlElement element = (XmlElement)nodeXml;
            string strID = element.GetAttribute("Item");
            Debug.WriteLine(strID);
            if (strID.CompareTo("Throughput Test Condition") != 0)
            {
                MessageBox.Show("This XML file is incorrect.", "Error");
                return false;
            }

            ///
            /// Read Function Test configuration settings
            ///            

            XmlNode nodeTestConditionModel = doc.SelectSingleNode("/CyberRouterATETC/TestCondition/Condition_Number");
            try
            {
                string Number = nodeTestConditionModel.SelectSingleNode("Number").InnerText;
                Debug.WriteLine("Number: " + Name);
                number = Int32.Parse(Number);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATETC/TestCondition/Condition_Number" + ex);
            }

            for (int i = 0; i < number; i++)
            {
                XmlNode nodeTestCondition = doc.SelectSingleNode("/CyberRouterATETC/TestCondition/Condition_" + (i + 1).ToString());

                try
                {
                    string Band = nodeTestCondition.SelectSingleNode("Band").InnerText;
                    string Mode = nodeTestCondition.SelectSingleNode("Mode").InnerText;
                    string SSID_Config = nodeTestCondition.SelectSingleNode("SSID_Config").InnerText;
                    string SSID_Text = nodeTestCondition.SelectSingleNode("SSID_Text").InnerText;
                    string Channel_Config = nodeTestCondition.SelectSingleNode("Channel_Config").InnerText;
                    string Channel = nodeTestCondition.SelectSingleNode("Channel").InnerText;                    
                    string Security_Config = nodeTestCondition.SelectSingleNode("Security_Config").InnerText;
                    string Security_Mode = nodeTestCondition.SelectSingleNode("Security_Mode").InnerText;
                    string Passphrase = nodeTestCondition.SelectSingleNode("Passphrase").InnerText;                    
                    string TX_tst_File = nodeTestCondition.SelectSingleNode("TX_tst_File").InnerText;
                    string RX_tst_File = nodeTestCondition.SelectSingleNode("RX_tst_File").InnerText;
                    string BiDir_tst_File = nodeTestCondition.SelectSingleNode("BiDir_tst_File").InnerText;
                 
                    Debug.WriteLine("Band: " + Band);
                    Debug.WriteLine("Mode: " + Mode);
                    Debug.WriteLine("SSID_Config?: " + SSID_Config);
                    Debug.WriteLine("SSID_Text: " + SSID_Text);
                    Debug.WriteLine("Channel_Config: " + Channel_Config);
                    Debug.WriteLine("Channel: " + Channel);                    
                    Debug.WriteLine("Security_Config: " + Security_Config);
                    Debug.WriteLine("Security_Mode: " + Security_Mode);
                    Debug.WriteLine("Passphrase: " + Passphrase);                    
                    Debug.WriteLine("TX_tst_File: " + TX_tst_File);
                    Debug.WriteLine("RX_tst_File: " + RX_tst_File);
                    Debug.WriteLine("BiDir_tst_File: " + BiDir_tst_File);

                    string[] data = new string[] { Band, Mode, SSID_Config, SSID_Text, Channel_Config, Channel, Security_Config, Security_Mode, Passphrase, TX_tst_File, RX_tst_File, BiDir_tst_File };
                    dgvThroughputTestConditionData.Rows.Add(data);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("/CyberRouterATETC/TestCondition/Condition_Read " + ex);
                }
            }

            // End of Read Test Condition configuration settings

            return true;
        }

       


    }
}
