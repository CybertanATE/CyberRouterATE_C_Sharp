///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RvRTurnTestCondition.cs
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
        /* Declare global variable */
        //int Channel_2_4G_Bound = 11 ;
        //int[] Channel_5G_20M = { 36, 40, 44, 48, 149, 153, 157, 161, 165};
        //int[] Channel_5G_40M = { 36, 40, 44, 48, 149, 153, 157, 161};
        //int[] Channel_5G_80M = { 36, 40, 44, 48, 149, 153, 157, 161};
        //string[] Mode_2_4G = {"20M", "40M"};
        //string[] Mode_5G = { "11N_20M", "11N_40M", "11AC_20M", "11AC_40M", "11AC_80M" };
        
        //ListBox gpib_IP;
        //int gpib_no;

        /* Declare variable to indicate Delete button exist in the datagridview1 column 0*/
        //bool hasDeleteButton = false;

        private void InitRvRTurnTestCondition()
        {
            rbtnRvRTurnTestCondition24G.Checked = true;
            txtRvRTurnTestConditionSsid.Text = "RvRTurnTest24";
            hasDeleteButton = false;

            cboxRvRTurnTestConditionWirelessSecurity.SelectedIndex = 0;
            txtRvRTurnTestConditionPassphrase.Text = "X";

            SetupDataGridViewRvRTurn();
            //SetupAttenuation();

            if (cboxRvRTurnTestConditionWirelessMode.SelectedIndex < 0)
            {
                cboxRvRTurnTestConditionWirelessChannel.Items.Clear();
                cboxRvRTurnTestConditionWirelessMode.Items.Clear();

                int i;
                for (i = 0; i < Mode_2_4G.Length; i++)
                {
                    cboxRvRTurnTestConditionWirelessMode.Items.Add(Mode_2_4G[i]);
                }

                cboxRvRTurnTestConditionWirelessMode.SelectedIndex = 0;


                for (i = 1; i < Channel_2_4G_Bound + 1; i++)
                {
                    cboxRvRTurnTestConditionWirelessChannel.Items.Add(i.ToString());
                }

                cboxRvRTurnTestConditionWirelessChannel.SelectedIndex = 0;
            }           
        }

        public void SetupDataGridViewRvRTurn()
        {
            dgvRvRTurnTestConditionData.ColumnCount = 12;
            dgvRvRTurnTestConditionData.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvRvRTurnTestConditionData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvRvRTurnTestConditionData.Name = "Test Condition Setting";
            //dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgvRvRTurnTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvRvRTurnTestConditionData.Columns[0].Name = "Band";
            dgvRvRTurnTestConditionData.Columns[1].Name = "Mode";
            dgvRvRTurnTestConditionData.Columns[2].Name = "Channel";
            dgvRvRTurnTestConditionData.Columns[3].Name = "SSID";
            dgvRvRTurnTestConditionData.Columns[4].Name = "Security";    
            dgvRvRTurnTestConditionData.Columns[5].Name = "Key";                     
            dgvRvRTurnTestConditionData.Columns[6].Name = "Start";
            dgvRvRTurnTestConditionData.Columns[7].Name = "Stop";
            dgvRvRTurnTestConditionData.Columns[8].Name = "Step";
            dgvRvRTurnTestConditionData.Columns[9].Name = "TX tst File";
            dgvRvRTurnTestConditionData.Columns[10].Name = "RX tst File";
            dgvRvRTurnTestConditionData.Columns[11].Name = "BI-Dir tst File";

            dgvRvRTurnTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvRvRTurnTestConditionData.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvRvRTurnTestConditionData.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False; //標題列換行, true -換行, false-不換行
            dgvRvRTurnTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//標題列置中

            dgvRvRTurnTestConditionData.Columns[0].Width = 80;
            dgvRvRTurnTestConditionData.Columns[1].Width = 80;
            dgvRvRTurnTestConditionData.Columns[2].Width = 80;
            dgvRvRTurnTestConditionData.Columns[3].Width = 120;
            dgvRvRTurnTestConditionData.Columns[4].Width = 80;
            dgvRvRTurnTestConditionData.Columns[5].Width = 80;
            dgvRvRTurnTestConditionData.Columns[6].Width = 80;
            dgvRvRTurnTestConditionData.Columns[7].Width = 80;
            dgvRvRTurnTestConditionData.Columns[8].Width = 80;
            dgvRvRTurnTestConditionData.Columns[9].Width = 120;
            dgvRvRTurnTestConditionData.Columns[10].Width = 120;
            dgvRvRTurnTestConditionData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dgvRvRTurnTestConditionData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dgvRvRTurnTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; //Use the column width setting
            dgvRvRTurnTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

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

        private void btnRvRTurnTestConditionAddSetting_Click(object sender, EventArgs e)
        {
            if (txtRvRTurnTestConditionSsid.Text == "") txtRvRTurnTestConditionSsid.Text = "RvRTurnTest";

            /* Check at least one tst file has selected */
            if (txtRvRTurnTestConditionTxTst.Text == "" && txtRvRTurnTestConditionRxTst.Text == "" && txtRvRTurnTestConditionBiTst.Text == "")
            {
                MessageBox.Show("Select at least one TST File!!", "Warning");
                return;
            }

            /* Check if attenuation stop value < start value */
            if (nudRvRTurnFunctionTestTurnTable1Stop.Value < nudRvRTurnFunctionTestTurnTable1Start.Value)
            {
                MessageBox.Show("Attenuation Stop Value must equal or greater than Start Value.", "Warning");
                return;
            }

            /* Check if Delete button exist in the datagridview */
            if (hasDeleteButton)
            {
                hasDeleteButton = false;
                dgvRvRTurnTestConditionData.Columns.RemoveAt(0);
                /* Reset the name of btn_TestCondition_Edit as Edit words. */
                btnRvRTurnTestConditionEditSetting.Text = "Edit";
            }

            //dgvRvRTurnTestConditionData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dgvRvRTurnTestConditionData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            /* Add data to datagridview */
            string[] row = new string[] {
                rbtnRvRTurnTestCondition24G.Checked ? "2.4G" : "5G", //Band
                cboxRvRTurnTestConditionWirelessMode.SelectedItem.ToString(), //Mode
                cboxRvRTurnTestConditionWirelessChannel.SelectedItem.ToString(), //Channel
                txtRvRTurnTestConditionSsid.Text,//SSID
                cboxRvRTurnTestConditionWirelessSecurity.SelectedItem.ToString(), //Security
                txtRvRTurnTestConditionPassphrase.Text, //Passphrase
                nudRvRTurnTestConditionAtteuationStart.Value.ToString(), //Attenuation start               
                nudRvRTurnTestConditionAtteuationStop.Value.ToString(), //Attenuation stop
                nudRvRTurnTestConditionAtteuationStep.Value.ToString(), //Attenuation step
                txtRvRTurnTestConditionTxTst.Text, //TX tst file
                txtRvRTurnTestConditionRxTst.Text, //RX tst File
                txtRvRTurnTestConditionBiTst.Text}; //Bi-Direction tst File*/};
            
            dgvRvRTurnTestConditionData.Rows.Add(row);
            dgvRvRTurnTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvRvRTurnTestConditionData.AutoResizeColumns();
            dgvRvRTurnTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
        }

        private void btnRvRTurnTestConditionEditSetting_Click(object sender, EventArgs e)
        {
            /* James: To prevent one cell is to be inserted more time when Edit button clicked once again. */
            if (hasDeleteButton == false)
            {
                hasDeleteButton = true;
                btnRvRTurnTestConditionEditSetting.Text = "Cancel";
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.Width = 66;
                btn.HeaderText = "Action";
                btn.Text = "Delete";
                btn.Name = "Action";
                btn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                btn.UseColumnTextForButtonValue = true;

                dgvRvRTurnTestConditionData.Columns.Insert(0, btn);
            }
            else
            {
                btnRvRTurnTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRvRTurnTestConditionData.Columns.Remove("Action");
            }
        }

        private void btnRvRTurnTestConditionSaveSetting_Click(object sender, EventArgs e)
        {
            if (dgvRvRTurnTestConditionData.RowCount <= 1)
            {
                MessageBox.Show("Add Condition First!!", "Warning");
                return;
            }

            string filename = "RvRTurnTestCondition";

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
                writeXmlRvRTurnTestCondition(saveFileDialog1.FileName);
            }
        }

        private void btnRvRTurnTestConditionLoadSetting_Click(object sender, EventArgs e)
        {
            if (btnRvRTurnTestConditionEditSetting.Text.ToLower() == "cancel")
            {
                btnRvRTurnTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRvRTurnTestConditionData.Columns.Remove("Action");
            }

            if (dgvRvRTurnTestConditionData.RowCount > 1)
            {
                if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                    return;
                else
                {
                    DataTable dt = (DataTable)dgvRvRTurnTestConditionData.DataSource;
                    dgvRvRTurnTestConditionData.Rows.Clear();
                    dgvRvRTurnTestConditionData.DataSource = dt;
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
                readXmlRvRTurnTestCondition(openFileDialog1.FileName);
            }
        }
        
        private void dgvRvRTurnTestConditionData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {        
            /* James modified: Adding a condition to prevent exception error when remove null cell. */
            if (dgvRvRTurnTestConditionData.Columns[e.ColumnIndex].Name == "Action" && hasDeleteButton == true && dgvRvRTurnTestConditionData.Rows[e.RowIndex].Cells[0].Value != null)
            {                
                dgvRvRTurnTestConditionData.Rows.RemoveAt(e.RowIndex);
                if (e.ColumnIndex == 0 && (String)dgvRvRTurnTestConditionData.Rows[0].Cells[0].Value == String.Empty)
                    dgvRvRTurnTestConditionData.Columns.RemoveAt(0);
            }
        }
                
        private void txtRvRTurnTestConditionTxTst_Click(object sender, EventArgs e)
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
                    txtRvRTurnTestConditionTxTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open TX TST file: " + ex.Message);
                }
            }    
        }

        private void txtRvRTurnTestConditionRxTst_Click(object sender, EventArgs e)
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
                    txtRvRTurnTestConditionRxTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open RX TST file: " + ex.Message);
                }
            }    
        }

        private void txtRvRTurnTestConditionBiTst_Click(object sender, EventArgs e)
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
                    txtRvRTurnTestConditionBiTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open Bi-Direction TST file: " + ex.Message);
                }
            }     
        }

        private void rbtnRvRTurnTestCondition24G_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnRvRTurnTestCondition24G.Checked == true)
            {        
                cboxRvRTurnTestConditionWirelessChannel.Items.Clear();
                cboxRvRTurnTestConditionWirelessMode.Items.Clear();

               int i;
                for(i=0 ;i< Mode_2_4G.Length ; i++)
                {
                    cboxRvRTurnTestConditionWirelessMode.Items.Add(Mode_2_4G[i]);
                }
                
                cboxRvRTurnTestConditionWirelessMode.SelectedIndex = 0;


                for (i = 1; i < Channel_2_4G_Bound+1; i++)
                {                   
                    cboxRvRTurnTestConditionWirelessChannel.Items.Add(i.ToString());
                }
                
                cboxRvRTurnTestConditionWirelessChannel.SelectedIndex = 0;
                
            }
        }

        private void rbtnRvRTurnTestCondition5G_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnRvRTurnTestCondition5G.Checked == true)
            {               
                cboxRvRTurnTestConditionWirelessChannel.Items.Clear();
                cboxRvRTurnTestConditionWirelessMode.Items.Clear();

                int i;
                for (i = 0; i < Mode_5G.Length; i++)
                {
                    cboxRvRTurnTestConditionWirelessMode.Items.Add(Mode_5G[i]);
                }
                cboxRvRTurnTestConditionWirelessMode.SelectedIndex = 0;

                //for (i = 0; i < Channel_5G_20M.Length; i++)
                //{
                //    cboxRvRTurnTestConditionWirelessChannel.Items.Add(Channel_5G_20M[i]);
                //}               
                //cboxRvRTurnTestConditionWirelessChannel.SelectedIndex = 0;
            }
        }

        private void cboxRvRTurnTestConditionWirelessMode_SelectedIndexChanged(object sender, EventArgs e)
        {       
            if (rbtnRvRTurnTestCondition5G.Checked == true)
            {
                //cboxRvRTurnTestConditionWirelessChannel.SelectedIndex = -1;
                cboxRvRTurnTestConditionWirelessChannel.Items.Clear();
                int i;

                if (cboxRvRTurnTestConditionWirelessMode.SelectedItem.ToString() == "11N_20M" ||                    
                    cboxRvRTurnTestConditionWirelessMode.SelectedItem.ToString() == "11AC_20M")
                {
                    for (i = 0; i < Channel_5G_20M.Length; i++)
                    {
                        cboxRvRTurnTestConditionWirelessChannel.Items.Add(Channel_5G_20M[i]);
                    }
                    cboxRvRTurnTestConditionWirelessChannel.SelectedIndex = 0;
                }
                    

                if (cboxRvRTurnTestConditionWirelessMode.SelectedItem.ToString() == "11N_40M" ||
                    cboxRvRTurnTestConditionWirelessMode.SelectedItem.ToString() == "11AC_40M")
                {
                    for (i = 0; i < Channel_5G_40M.Length; i++)
                    {
                        cboxRvRTurnTestConditionWirelessChannel.Items.Add(Channel_5G_40M[i]);
                    }
                    cboxRvRTurnTestConditionWirelessChannel.SelectedIndex = 0;

                }
                if (cboxRvRTurnTestConditionWirelessMode.SelectedItem.ToString() == "11AC_80M")
                {
                    for (i = 0; i < Channel_5G_80M.Length; i++)
                    {
                        cboxRvRTurnTestConditionWirelessChannel.Items.Add(Channel_5G_80M[i]);
                    }
                    cboxRvRTurnTestConditionWirelessChannel.SelectedIndex = 0;
                }                
            }
        }

        public void writeXmlRvRTurnTestCondition(string FileName)
        {
            // ToDo: Needs to verify which test item is to be selected..

            string[,] rowdata = new string[dgvRvRTurnTestConditionData.RowCount - 1, 7];

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;

            XmlWriter writer = XmlWriter.Create(FileName, settings);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by the program.");
            writer.WriteStartElement("CyberRouterATETC");
            writer.WriteAttributeString("Item", "RvRTurn Test Condition");

            ///
            /// Write Function Test settings
            /// 
            writer.WriteStartElement("TestCondition");
            for (int i = 0; i < dgvRvRTurnTestConditionData.RowCount - 1; i++)
            {
                writer.WriteStartElement("Condition_" + (i + 1).ToString());

                writer.WriteElementString("Band", rowdata[i, 0] = dgvRvRTurnTestConditionData.Rows[i].Cells[0].Value.ToString());
                writer.WriteElementString("Mode", rowdata[i, 1] = dgvRvRTurnTestConditionData.Rows[i].Cells[1].Value.ToString());
                writer.WriteElementString("Channel", rowdata[i, 2] = dgvRvRTurnTestConditionData.Rows[i].Cells[2].Value.ToString());
                writer.WriteElementString("SSID", rowdata[i, 3] = dgvRvRTurnTestConditionData.Rows[i].Cells[3].Value.ToString());
                writer.WriteElementString("Security", rowdata[i, 3] = dgvRvRTurnTestConditionData.Rows[i].Cells[4].Value.ToString());
                writer.WriteElementString("Key", rowdata[i, 3] = dgvRvRTurnTestConditionData.Rows[i].Cells[5].Value.ToString());                
                writer.WriteElementString("Start", rowdata[i, 4] = dgvRvRTurnTestConditionData.Rows[i].Cells[6].Value.ToString());
                writer.WriteElementString("Stop", rowdata[i, 5] = dgvRvRTurnTestConditionData.Rows[i].Cells[7].Value.ToString());
                writer.WriteElementString("Step", rowdata[i, 6] = dgvRvRTurnTestConditionData.Rows[i].Cells[8].Value.ToString());
                writer.WriteElementString("TX_tst_File", rowdata[i, 6] = dgvRvRTurnTestConditionData.Rows[i].Cells[9].Value.ToString());
                writer.WriteElementString("RX_tst_File", rowdata[i, 6] = dgvRvRTurnTestConditionData.Rows[i].Cells[10].Value.ToString());
                writer.WriteElementString("BI_Dir_tst_File", rowdata[i, 6] = dgvRvRTurnTestConditionData.Rows[i].Cells[11].Value.ToString());

                writer.WriteEndElement();
            }

            writer.WriteStartElement("Condition_Number");
            writer.WriteElementString("Number", dgvRvRTurnTestConditionData.RowCount.ToString());
            writer.WriteEndElement();

            // End of write Function Test
            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public bool readXmlRvRTurnTestCondition(string FileName)
        {
            int number = 0;

            dgvRvRTurnTestConditionData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            XmlDocument doc = new XmlDocument();
            doc.Load(FileName);

            XmlNode nodeXml = doc.SelectSingleNode("CyberRouterATETC");
            if (nodeXml == null)
                return false;
            XmlElement element = (XmlElement)nodeXml;
            string strID = element.GetAttribute("Item");
            Debug.WriteLine(strID);
            if (strID.CompareTo("RvRTurn Test Condition") != 0)
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

                    dgvRvRTurnTestConditionData.Columns[3].Name = "Security";
                    dgvRvRTurnTestConditionData.Columns[4].Name = "Key";

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
                    dgvRvRTurnTestConditionData.Rows.Add(data);
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
