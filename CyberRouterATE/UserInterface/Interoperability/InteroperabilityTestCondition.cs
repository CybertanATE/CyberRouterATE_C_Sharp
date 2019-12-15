///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : InteroperabilityTestCondition.cs
///  Update         : 2015-05-11
///  Description    : Main function
///  Modified       : 2015-05-11 Initial version  
///                   
///---------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Xml;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        private void InitInteroperabilityTestCondition()
        {
            rbtnInteroperabilityTestCondition24G.Checked = true;
            txtInteroperabilityTestConditionSsid.Text = "InteroperabilityTest24";
            hasDeleteButton = false;

            cboxInteroperabilityTestConditionWirelessSecurity.SelectedIndex = 0;
            txtInteroperabilityTestConditionPassphrase.Text = "X";

            SetupDataGridViewInteroperability();
            //SetupAttenuation();

            if (cboxInteroperabilityTestConditionWirelessMode.SelectedIndex < 0)
            {
                cboxInteroperabilityTestConditionWirelessChannel.Items.Clear();
                cboxInteroperabilityTestConditionWirelessMode.Items.Clear();

                int i;
                for (i = 0; i < Mode_2_4G.Length; i++)
                {
                    cboxInteroperabilityTestConditionWirelessMode.Items.Add(Mode_2_4G[i]);
                }

                cboxInteroperabilityTestConditionWirelessMode.SelectedIndex = 0;


                for (i = 1; i < Channel_2_4G_Bound + 1; i++)
                {
                    cboxInteroperabilityTestConditionWirelessChannel.Items.Add(i.ToString());
                }

                cboxInteroperabilityTestConditionWirelessChannel.SelectedIndex = 0;
            }           
        }

        public void SetupDataGridViewInteroperability()
        {
            dgvInteroperabilityTestConditionData.ColumnCount = 9;
            dgvInteroperabilityTestConditionData.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvInteroperabilityTestConditionData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvInteroperabilityTestConditionData.Name = "Test Condition Setting";
            //dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgvInteroperabilityTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvInteroperabilityTestConditionData.Columns[0].Name = "Band";
            dgvInteroperabilityTestConditionData.Columns[1].Name = "Mode";
            dgvInteroperabilityTestConditionData.Columns[2].Name = "Channel";
            dgvInteroperabilityTestConditionData.Columns[3].Name = "SSID";
            dgvInteroperabilityTestConditionData.Columns[4].Name = "Security";    
            dgvInteroperabilityTestConditionData.Columns[5].Name = "Key";            
            dgvInteroperabilityTestConditionData.Columns[6].Name = "TX tst File";
            dgvInteroperabilityTestConditionData.Columns[7].Name = "RX tst File";
            dgvInteroperabilityTestConditionData.Columns[8].Name = "BI-Dir tst File";

            dgvInteroperabilityTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvInteroperabilityTestConditionData.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvInteroperabilityTestConditionData.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False; //標題列換行, true -換行, false-不換行
            dgvInteroperabilityTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//標題列置中

            dgvInteroperabilityTestConditionData.Columns[0].Width = 80;
            dgvInteroperabilityTestConditionData.Columns[1].Width = 80;
            dgvInteroperabilityTestConditionData.Columns[2].Width = 80;
            dgvInteroperabilityTestConditionData.Columns[3].Width = 120;
            dgvInteroperabilityTestConditionData.Columns[4].Width = 80;
            dgvInteroperabilityTestConditionData.Columns[5].Width = 80;            
            dgvInteroperabilityTestConditionData.Columns[6].Width = 120;
            dgvInteroperabilityTestConditionData.Columns[7].Width = 120;
            dgvInteroperabilityTestConditionData.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dgvInteroperabilityTestConditionData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dgvInteroperabilityTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; //Use the column width setting
            dgvInteroperabilityTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            /*
            string[] row = { (dataGridView1.NewRowIndex + 1).ToString(), numericUpDown1.Value.ToString(), numericUpDown2.Value.ToString(), numericUpDown3.Value.ToString(), textBox1.Text, textBox2.Text };
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoResizeColumns();

            //dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            //dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.Rows.Add(row);
            dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows.Count - 1;
            */
        }
              
        private void btnInteroperabilityTestConditionAddSetting_Click(object sender, EventArgs e)
        {
            if (txtInteroperabilityTestConditionSsid.Text == "") txtInteroperabilityTestConditionSsid.Text = "InteroperabilityTest";

            /* Check at least one tst file has selected */
            if (txtInteroperabilityTestConditionTxTst.Text == "" && txtInteroperabilityTestConditionRxTst.Text == "" && txtInteroperabilityTestConditionBiTst.Text == "")
            {
                MessageBox.Show("Select at least one TST File!!", "Warning");
                return;
            }                      

            /* Check if Delete button exist in the datagridview */
            if (hasDeleteButton)
            {
                hasDeleteButton = false;
                dgvInteroperabilityTestConditionData.Columns.RemoveAt(0);
                /* Reset the name of btn_TestCondition_Edit as Edit words. */
                btnInteroperabilityTestConditionEditSetting.Text = "Edit";
            }

            //dgvInteroperabilityTestConditionData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dgvInteroperabilityTestConditionData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            /* Add data to datagridview */
            string[] row = new string[] {
                rbtnInteroperabilityTestCondition24G.Checked ? "2.4G" : "5G", //Band
                cboxInteroperabilityTestConditionWirelessMode.SelectedItem.ToString(), //Mode
                cboxInteroperabilityTestConditionWirelessChannel.SelectedItem.ToString(), //Channel
                txtInteroperabilityTestConditionSsid.Text,//SSID
                cboxInteroperabilityTestConditionWirelessSecurity.SelectedItem.ToString(), //Security
                txtInteroperabilityTestConditionPassphrase.Text, //Passphrase                
                txtInteroperabilityTestConditionTxTst.Text, //TX tst file
                txtInteroperabilityTestConditionRxTst.Text, //RX tst File
                txtInteroperabilityTestConditionBiTst.Text}; //Bi-Direction tst File*/};
            
            dgvInteroperabilityTestConditionData.Rows.Add(row);
            dgvInteroperabilityTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvInteroperabilityTestConditionData.AutoResizeColumns();
            dgvInteroperabilityTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
        }

        private void btnInteroperabilityTestConditionEditSetting_Click(object sender, EventArgs e)       
        {
            /* James: To prevent one cell is to be inserted more time when Edit button clicked once again. */
            if (hasDeleteButton == false)
            {
                hasDeleteButton = true;
                btnInteroperabilityTestConditionEditSetting.Text = "Cancel";
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.Width = 66;
                btn.HeaderText = "Action";
                btn.Text = "Delete";
                btn.Name = "Action";
                btn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                btn.UseColumnTextForButtonValue = true;

                dgvInteroperabilityTestConditionData.Columns.Insert(0, btn);
            }
            else
            {
                btnInteroperabilityTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvInteroperabilityTestConditionData.Columns.Remove("Action");
            }
        }

        private void btnInteroperabilityTestConditionSaveSetting_Click(object sender, EventArgs e)        
        {
            if (dgvInteroperabilityTestConditionData.RowCount <= 1)
            {
                MessageBox.Show("Add Condition First!!", "Warning");
                return;
            }

            string filename = "InteroperabilityTestCondition";

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
                writeXmlInteroperabilityTestCondition(saveFileDialog1.FileName);
            }
        }

        private void btnInteroperabilityTestConditionLoadSetting_Click(object sender, EventArgs e)        
        {
            if (btnInteroperabilityTestConditionEditSetting.Text.ToLower() == "cancel")
            {
                btnInteroperabilityTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvInteroperabilityTestConditionData.Columns.Remove("Action");
            }

            if (dgvInteroperabilityTestConditionData.RowCount > 1)
            {
                if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                    return;
                else
                {
                    DataTable dt = (DataTable)dgvInteroperabilityTestConditionData.DataSource;
                    dgvInteroperabilityTestConditionData.Rows.Clear();
                    dgvInteroperabilityTestConditionData.DataSource = dt;
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
                readXmlInteroperabilityTestCondition(openFileDialog1.FileName);
            }
        }
        
        private void dgvInteroperabilityTestConditionData_CellContentClick(object sender, DataGridViewCellEventArgs e)        
        {        
            /* James modified: Adding a condition to prevent exception error when remove null cell. */
            if (dgvInteroperabilityTestConditionData.Columns[e.ColumnIndex].Name == "Action" && hasDeleteButton == true && dgvInteroperabilityTestConditionData.Rows[e.RowIndex].Cells[0].Value != null)
            {                
                dgvInteroperabilityTestConditionData.Rows.RemoveAt(e.RowIndex);
                if (e.ColumnIndex == 0 && (String)dgvInteroperabilityTestConditionData.Rows[0].Cells[0].Value == String.Empty)
                    dgvInteroperabilityTestConditionData.Columns.RemoveAt(0);
            }
        }
        
        private void txtInteroperabilityTestConditionTxTst_Click(object sender, EventArgs e)        
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
                    txtInteroperabilityTestConditionTxTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open TX TST file: " + ex.Message);
                }
            }    
        }

        private void txtInteroperabilityTestConditionRxTst_Click(object sender, EventArgs e)        
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
                    txtInteroperabilityTestConditionRxTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open RX TST file: " + ex.Message);
                }
            }    
        }

        private void txtInteroperabilityTestConditionBiTst_Click(object sender, EventArgs e)        
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
                    txtInteroperabilityTestConditionBiTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open Bi-Direction TST file: " + ex.Message);
                }
            }     
        }
        
        private void rbtnInteroperabilityTestCondition24G_CheckedChanged(object sender, EventArgs e)        
        {
            if (rbtnInteroperabilityTestCondition24G.Checked == true)
            {        
                cboxInteroperabilityTestConditionWirelessChannel.Items.Clear();
                cboxInteroperabilityTestConditionWirelessMode.Items.Clear();

               int i;
                for(i=0 ;i< Mode_2_4G.Length ; i++)
                {
                    cboxInteroperabilityTestConditionWirelessMode.Items.Add(Mode_2_4G[i]);
                }
                
                cboxInteroperabilityTestConditionWirelessMode.SelectedIndex = 0;


                for (i = 1; i < Channel_2_4G_Bound+1; i++)
                {                   
                    cboxInteroperabilityTestConditionWirelessChannel.Items.Add(i.ToString());
                }
                
                cboxInteroperabilityTestConditionWirelessChannel.SelectedIndex = 0;
                
            }
        }

        private void rbtnInteroperabilityTestCondition5G_CheckedChanged(object sender, EventArgs e)        
        {
            if (rbtnInteroperabilityTestCondition5G.Checked == true)
            {               
                cboxInteroperabilityTestConditionWirelessChannel.Items.Clear();
                cboxInteroperabilityTestConditionWirelessMode.Items.Clear();

                int i;
                for (i = 0; i < Mode_5G.Length; i++)
                {
                    cboxInteroperabilityTestConditionWirelessMode.Items.Add(Mode_5G[i]);
                }
                cboxInteroperabilityTestConditionWirelessMode.SelectedIndex = 0;

                //for (i = 0; i < Channel_5G_20M.Length; i++)
                //{
                //    cboxInteroperabilityTestConditionWirelessChannel.Items.Add(Channel_5G_20M[i]);
                //}               
                //cboxInteroperabilityTestConditionWirelessChannel.SelectedIndex = 0;
            }
        }

        private void cboxInteroperabilityTestConditionWirelessMode_SelectedIndexChanged(object sender, EventArgs e)        
        {       
            if (rbtnInteroperabilityTestCondition5G.Checked == true)
            {
                //cboxInteroperabilityTestConditionWirelessChannel.SelectedIndex = -1;
                cboxInteroperabilityTestConditionWirelessChannel.Items.Clear();
                int i;

                if (cboxInteroperabilityTestConditionWirelessMode.SelectedItem.ToString() == "11N_20M" ||                    
                    cboxInteroperabilityTestConditionWirelessMode.SelectedItem.ToString() == "11AC_20M")
                {
                    for (i = 0; i < Channel_5G_20M.Length; i++)
                    {
                        cboxInteroperabilityTestConditionWirelessChannel.Items.Add(Channel_5G_20M[i]);
                    }
                    cboxInteroperabilityTestConditionWirelessChannel.SelectedIndex = 0;
                }
                    

                if (cboxInteroperabilityTestConditionWirelessMode.SelectedItem.ToString() == "11N_40M" ||
                    cboxInteroperabilityTestConditionWirelessMode.SelectedItem.ToString() == "11AC_40M")
                {
                    for (i = 0; i < Channel_5G_40M.Length; i++)
                    {
                        cboxInteroperabilityTestConditionWirelessChannel.Items.Add(Channel_5G_40M[i]);
                    }
                    cboxInteroperabilityTestConditionWirelessChannel.SelectedIndex = 0;

                }
                if (cboxInteroperabilityTestConditionWirelessMode.SelectedItem.ToString() == "11AC_80M")
                {
                    for (i = 0; i < Channel_5G_80M.Length; i++)
                    {
                        cboxInteroperabilityTestConditionWirelessChannel.Items.Add(Channel_5G_80M[i]);
                    }
                    cboxInteroperabilityTestConditionWirelessChannel.SelectedIndex = 0;
                }                
            }
        }

        public void writeXmlInteroperabilityTestCondition(string FileName)
        {
            // ToDo: Needs to verify which test item is to be selected..

            string[,] rowdata = new string[dgvInteroperabilityTestConditionData.RowCount - 1, 7];

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;

            XmlWriter writer = XmlWriter.Create(FileName, settings);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by the program.");
            writer.WriteStartElement("CyberRouterATETC");
            writer.WriteAttributeString("Item", "Interoperability Test Condition");

            ///
            /// Write Function Test settings
            /// 
            writer.WriteStartElement("TestCondition");
            for (int i = 0; i < dgvInteroperabilityTestConditionData.RowCount - 1; i++)
            {
                writer.WriteStartElement("Condition_" + (i + 1).ToString());

                writer.WriteElementString("Band", rowdata[i, 0] = dgvInteroperabilityTestConditionData.Rows[i].Cells[0].Value.ToString());
                writer.WriteElementString("Mode", rowdata[i, 1] = dgvInteroperabilityTestConditionData.Rows[i].Cells[1].Value.ToString());
                writer.WriteElementString("Channel", rowdata[i, 2] = dgvInteroperabilityTestConditionData.Rows[i].Cells[2].Value.ToString());
                writer.WriteElementString("SSID", rowdata[i, 3] = dgvInteroperabilityTestConditionData.Rows[i].Cells[3].Value.ToString());
                writer.WriteElementString("Security", rowdata[i, 3] = dgvInteroperabilityTestConditionData.Rows[i].Cells[4].Value.ToString());
                writer.WriteElementString("Key", rowdata[i, 3] = dgvInteroperabilityTestConditionData.Rows[i].Cells[5].Value.ToString());                
                writer.WriteElementString("Start", rowdata[i, 4] = dgvInteroperabilityTestConditionData.Rows[i].Cells[6].Value.ToString());
                writer.WriteElementString("Stop", rowdata[i, 5] = dgvInteroperabilityTestConditionData.Rows[i].Cells[7].Value.ToString());
                writer.WriteElementString("Step", rowdata[i, 6] = dgvInteroperabilityTestConditionData.Rows[i].Cells[8].Value.ToString());
                writer.WriteElementString("TX_tst_File", rowdata[i, 6] = dgvInteroperabilityTestConditionData.Rows[i].Cells[9].Value.ToString());
                writer.WriteElementString("RX_tst_File", rowdata[i, 6] = dgvInteroperabilityTestConditionData.Rows[i].Cells[10].Value.ToString());
                writer.WriteElementString("BI_Dir_tst_File", rowdata[i, 6] = dgvInteroperabilityTestConditionData.Rows[i].Cells[11].Value.ToString());

                writer.WriteEndElement();
            }

            writer.WriteStartElement("Condition_Number");
            writer.WriteElementString("Number", dgvInteroperabilityTestConditionData.RowCount.ToString());
            writer.WriteEndElement();

            // End of write Function Test
            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public bool readXmlInteroperabilityTestCondition(string FileName)
        {
            int number = 0;

            dgvInteroperabilityTestConditionData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            XmlDocument doc = new XmlDocument();
            doc.Load(FileName);

            XmlNode nodeXml = doc.SelectSingleNode("CyberRouterATETC");
            if (nodeXml == null)
                return false;
            XmlElement element = (XmlElement)nodeXml;
            string strID = element.GetAttribute("Item");
            Debug.WriteLine(strID);
            if (strID.CompareTo("Interoperability Test Condition") != 0)
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
                    string Channel = nodeTestCondition.SelectSingleNode("Channel").InnerText;
                    string SSID = nodeTestCondition.SelectSingleNode("SSID").InnerText;
                    string Security = nodeTestCondition.SelectSingleNode("Security").InnerText;
                    string Key = nodeTestCondition.SelectSingleNode("Key").InnerText;
                    string Start = nodeTestCondition.SelectSingleNode("Start").InnerText;
                    string Stop = nodeTestCondition.SelectSingleNode("Stop").InnerText;
                    string Step = nodeTestCondition.SelectSingleNode("Step").InnerText;
                    string TX_tst_File = nodeTestCondition.SelectSingleNode("TX_tst_File").InnerText;
                    string RX_tst_File = nodeTestCondition.SelectSingleNode("RX_tst_File").InnerText;
                    string BI_Dir_tst_File = nodeTestCondition.SelectSingleNode("BI_Dir_tst_File").InnerText;

                    dgvInteroperabilityTestConditionData.Columns[3].Name = "Security";
                    dgvInteroperabilityTestConditionData.Columns[4].Name = "Key";

                    Debug.WriteLine("Band: " + Band);
                    Debug.WriteLine("Mode: " + Mode);
                    Debug.WriteLine("Channel: " + Channel);
                    Debug.WriteLine("SSID: " + SSID);
                    Debug.WriteLine("Security: " + Security);
                    Debug.WriteLine("Key: " + Key);
                    Debug.WriteLine("Start: " + Start);
                    Debug.WriteLine("Stop: " + Stop);
                    Debug.WriteLine("Step: " + Step);
                    Debug.WriteLine("TX_tst_File: " + TX_tst_File);
                    Debug.WriteLine("RX_tst_File: " + RX_tst_File);
                    Debug.WriteLine("BI_Dir_tst_File: " + BI_Dir_tst_File);

                    string[] data = new string[] { Band, Mode, Channel, SSID, Security, Key, Start, Stop, Step, TX_tst_File, RX_tst_File, BI_Dir_tst_File};
                    dgvInteroperabilityTestConditionData.Rows.Add(data);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("/CyberRouterATETC/TestCondition/Condition_Read " + ex);
                }
            }

            // End of Read Test Condition configuration settings

            return true;
        }




       








        //private void DefaultAttenuation()
        //{
        //    Attenuation_buttonValue_2_4G = new decimal[] { 1, 2, 4, 4, 10, 20, 30, 30 };
        //    Attenuation_buttonValue_5G = new decimal[] { 1, 2, 4, 4, 10, 20, 40, 40 };     
        //}
    }
}
