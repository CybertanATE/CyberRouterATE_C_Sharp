///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RvRFunctionTest.cs
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
        
        
        ListBox gpib_IP;
        int gpib_no;

        /* Declare variable to indicate Delete button exist in the datagridview1 column 0*/
        bool hasDeleteButton = false;

        private void InitRvRTestCondition()
        {
            rbtnRvRTestCondition24G.Checked = true;
            txt_TestCondition_SSID.Text = "RvRTest24";
            hasDeleteButton = false;

            SetupDataGridView();
            //SetupAttenuation();

            /* Add 2.4G mode 20M, 40M */
            /*
            dud_TestCondition_Mode.Items.Add("20M");
            dud_TestCondition_Mode.Items.Add("40M");
            dud_TestCondition_Mode.SelectedIndex = 0;
            cbox_TestCondition_Mode.Items.Add("20M");
            cbox_TestCondition_Mode.Items.Add("40M");
            cbox_TestCondition_Mode.SelectedIndex = 0;
            */

            /* Add 2.4G channel 1-13 */
            /*
            for (int i = 1; i < Channel_2_4G_Bound+1; i++)
            {
                dud_TestCondition_Channel.Items.Add(i.ToString());
                cbox_TestCondition_Channel.Items.Add(i.ToString());
            }
            dud_TestCondition_Channel.SelectedIndex = 0;
            cbox_TestCondition_Channel.SelectedIndex = 0;
            */
        }

        public void SetupDataGridView()
        {
            dgvRvRTestConditionData.ColumnCount = 10;
            dgvRvRTestConditionData.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvRvRTestConditionData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvRvRTestConditionData.Name = "Test Condition Setting";
            //dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgvRvRTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvRvRTestConditionData.Columns[0].Name = "Band";
            dgvRvRTestConditionData.Columns[1].Name = "Mode";
            dgvRvRTestConditionData.Columns[2].Name = "Channel";
            dgvRvRTestConditionData.Columns[3].Name = "SSID";
            dgvRvRTestConditionData.Columns[4].Name = "Start";
            dgvRvRTestConditionData.Columns[5].Name = "Stop";
            dgvRvRTestConditionData.Columns[6].Name = "Step";
            dgvRvRTestConditionData.Columns[7].Name = "TX tst File";
            dgvRvRTestConditionData.Columns[8].Name = "RX tst File";
            dgvRvRTestConditionData.Columns[9].Name = "BI-Dir tst File";

            dgvRvRTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvRvRTestConditionData.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvRvRTestConditionData.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False; //標題列換行, true -換行, false-不換行
            dgvRvRTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//標題列置中

            dgvRvRTestConditionData.Columns[0].Width = 80;
            dgvRvRTestConditionData.Columns[1].Width = 80;
            dgvRvRTestConditionData.Columns[2].Width = 80;
            dgvRvRTestConditionData.Columns[3].Width = 120;
            dgvRvRTestConditionData.Columns[4].Width = 80;
            dgvRvRTestConditionData.Columns[5].Width = 80;
            dgvRvRTestConditionData.Columns[6].Width = 80;
            dgvRvRTestConditionData.Columns[7].Width = 120;
            dgvRvRTestConditionData.Columns[8].Width = 120;
            dgvRvRTestConditionData.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dgvRvRTestConditionData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dgvRvRTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; //Use the column width setting
            dgvRvRTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

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

        private void btnRvRTestConditionAddSetting_Click(object sender, EventArgs e)
        {                   
            /* Check all data has value */
            if (txt_TestCondition_SSID.Text == "") txt_TestCondition_SSID.Text = "RvRTest";

            /* Check at least one tst file has selected */
            if (txtRvRTestConditionTxTst.Text == "" && txtRvRTestConditionRxTst.Text == "" && txtRvRTestConditionBiTst.Text == "")
            {
                MessageBox.Show("Select at least one TST File!!", "Warning");
                return;
            }

            /* Check if attenuation stop value < start value */
            if(nud_TestCondition_AtteuationMax.Value < nud_TestCondition_AtteuationMin.Value)
            {
                MessageBox.Show("Attenuation Stop Value must equal or greater than Start Value.", "Warning");
                return ;
            }

            /* Check if Delete button exist in the datagridview */
            if (hasDeleteButton)
            {
                hasDeleteButton = false;
                dgvRvRTestConditionData.Columns.RemoveAt(0);
                /* Reset the name of btn_TestCondition_Edit as Edit words. */
                btnRvRTestConditionEditSetting.Text = "Edit";
            }

            /* Add data to datagridview */
            string a1 = rbtnRvRTestCondition24G.Checked? "2.4G":"5G"; //Band
            string a2 = cboxRvRTestConditionWirelessMode.SelectedItem.ToString(); //Mode
            string a3 = cboxRvRTestConditionWirelessChannel.SelectedItem.ToString(); //Channel
            string a4 = txt_TestCondition_SSID.Text;//SSID
            string a5 = nud_TestCondition_AtteuationMin.Value.ToString(); //Attenuation min
            string a6 = nud_TestCondition_AtteuationMax.Value.ToString() ; //Attenuation max
            string a7 = nud_TestCondition_AtteuationStep.Value.ToString() ; //Attenuation step
            string a8 = txtRvRTestConditionTxTst.Text; //TX tst file
            string a9 = txtRvRTestConditionRxTst.Text; //RX tst File
            string a10 = txtRvRTestConditionBiTst.Text; //Bi-Direction tst File*/

            string[] row = {a1, a2, a3, a4, a5, a6, a7, a8, a9, a10 };            
            dgvRvRTestConditionData.Rows.Add(row);
            dgvRvRTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvRvRTestConditionData.AutoResizeColumns();
            dgvRvRTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;        
        }

        private void btnRvRTestConditionEdit_Click(object sender, EventArgs e)
        {       
            /* James: To prevent one cell is to be inserted more time when Edit button clicked once again. */
            if (hasDeleteButton == false)
            {
                hasDeleteButton = true;
                btnRvRTestConditionEditSetting.Text = "Cancel";
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.Width = 66;
                btn.HeaderText = "Action";
                btn.Text = "Delete";
                btn.Name = "Action";
                btn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                btn.UseColumnTextForButtonValue = true;

                dgvRvRTestConditionData.Columns.Insert(0, btn);
            }
            else
            {
                btnRvRTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRvRTestConditionData.Columns.Remove("Action");
            }
        }

        private void dgvRvRTestConditionData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {        
            /* James modified: Adding a condition to prevent exception error when remove null cell. */
            if (dgvRvRTestConditionData.Columns[e.ColumnIndex].Name == "Action" && hasDeleteButton == true && dgvRvRTestConditionData.Rows[e.RowIndex].Cells[0].Value != null)
            {                
                dgvRvRTestConditionData.Rows.RemoveAt(e.RowIndex);
                if (e.ColumnIndex == 0 && (String)dgvRvRTestConditionData.Rows[0].Cells[0].Value == String.Empty)
                    dgvRvRTestConditionData.Columns.RemoveAt(0);
            }
        }

        private void btnRvRTestConditionSaveSetting_Click(object sender, EventArgs e)
        {
            if (dgvRvRTestConditionData.RowCount <= 1)
            {
                MessageBox.Show("Add Condition First!!", "Warning");
                return;
            }

            string filename = "RvRTestCondition";

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
                writeXmlRvRTestCondition(saveFileDialog1.FileName);
            }

        }

        private void btnRvRTestConditionLoadSetting_Click(object sender, EventArgs e)
        {
            if (btnRvRTestConditionEditSetting.Text.ToLower() == "cancel")
            {
                btnRvRTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRvRTestConditionData.Columns.Remove("Action");
            }

            if (dgvRvRTestConditionData.RowCount > 1)
            {
                if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                    return;
                else
                {
                    DataTable dt = (DataTable)dgvRvRTestConditionData.DataSource;
                    dgvRvRTestConditionData.Rows.Clear();
                    dgvRvRTestConditionData.DataSource = dt;
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
                readXmlRvRTestCondition(openFileDialog1.FileName);
            }
        }
        
        private void txtRvRTestConditionTxTst_Click(object sender, EventArgs e)
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
                    txtRvRTestConditionTxTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open TX TST file: " + ex.Message);
                }
            }    
        }

        private void txtRvRTestConditionRxTst_Click(object sender, EventArgs e)
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
                    txtRvRTestConditionRxTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open RX TST file: " + ex.Message);
                }
            }    
        }

        private void txtRvRTestConditionBiTst_Click(object sender, EventArgs e)
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
                    txtRvRTestConditionBiTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open Bi-Direction TST file: " + ex.Message);
                }
            }     
        }

        private void rbtnRvRTestCondition24G_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnRvRTestCondition24G.Checked == true)
            {        
                cboxRvRTestConditionWirelessChannel.Items.Clear();
                cboxRvRTestConditionWirelessMode.Items.Clear();

               int i;
                for(i=0 ;i< Mode_2_4G.Length ; i++)
                {
                    cboxRvRTestConditionWirelessMode.Items.Add(Mode_2_4G[i]);
                }
                
                cboxRvRTestConditionWirelessMode.SelectedIndex = 0;


                for (i = 1; i < Channel_2_4G_Bound+1; i++)
                {                   
                    cboxRvRTestConditionWirelessChannel.Items.Add(i.ToString());
                }
                
                cboxRvRTestConditionWirelessChannel.SelectedIndex = 0;
                
            }
        }

        private void rbtnRvRTestCondition5G_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnRvRTestCondition5G.Checked == true)
            {               
                cboxRvRTestConditionWirelessChannel.Items.Clear();
                cboxRvRTestConditionWirelessMode.Items.Clear();

                int i;
                for (i = 0; i < Mode_5G.Length; i++)
                {
                    cboxRvRTestConditionWirelessMode.Items.Add(Mode_5G[i]);
                }
                cboxRvRTestConditionWirelessMode.SelectedIndex = 0;

                for (i = 0; i < Channel_5G_20M.Length; i++)
                {
                    cboxRvRTestConditionWirelessChannel.Items.Add(Channel_5G_20M[i]);
                }               
                cboxRvRTestConditionWirelessChannel.SelectedIndex = 0;
            }
        }

        private void cboxRvRTestConditionWirelessMode_SelectedIndexChanged(object sender, EventArgs e)
        {       
            if (rbtnRvRTestCondition5G.Checked == true)
            {
                cboxRvRTestConditionWirelessChannel.Items.Clear();
                int i;

                if (cboxRvRTestConditionWirelessMode.SelectedItem.ToString() == "11N_20M" ||                    
                    cboxRvRTestConditionWirelessMode.SelectedItem.ToString() == "11AC_20M")
                {
                    for (i = 0; i < Channel_5G_20M.Length; i++)
                    {
                        cboxRvRTestConditionWirelessChannel.Items.Add(Channel_5G_20M[i]);
                    }
                    cboxRvRTestConditionWirelessChannel.SelectedIndex = 0;
                }
                    

                if (cboxRvRTestConditionWirelessMode.SelectedItem.ToString() == "11N_40M" ||
                    cboxRvRTestConditionWirelessMode.SelectedItem.ToString() == "11AC_40M")
                {
                    for (i = 0; i < Channel_5G_40M.Length; i++)
                    {
                        cboxRvRTestConditionWirelessChannel.Items.Add(Channel_5G_40M[i]);
                    }
                    cboxRvRTestConditionWirelessChannel.SelectedIndex = 0;

                }
                if (cboxRvRTestConditionWirelessMode.SelectedItem.ToString() == "11AC_80M")
                {
                    for (i = 0; i < Channel_5G_80M.Length; i++)
                    {
                        cboxRvRTestConditionWirelessChannel.Items.Add(Channel_5G_80M[i]);
                    }
                    cboxRvRTestConditionWirelessChannel.SelectedIndex = 0;
                }                
            }
        }

        public void writeXmlRvRTestCondition(string FileName)
        {
            // ToDo: Needs to verify which test item is to be selected..

            string[,] rowdata = new string[dgvRvRTestConditionData.RowCount - 1, 7];

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;

            XmlWriter writer = XmlWriter.Create(FileName, settings);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by the program.");
            writer.WriteStartElement("CyberRouterATETC");
            writer.WriteAttributeString("Item", "RvR Test Condition");

            ///
            /// Write Function Test settings
            /// 
            writer.WriteStartElement("TestCondition");
            for (int i = 0; i < dgvRvRTestConditionData.RowCount - 1; i++)
            {
                writer.WriteStartElement("Condition_" + (i + 1).ToString());

                writer.WriteElementString("Band", rowdata[i, 0] = dgvRvRTestConditionData.Rows[i].Cells[0].Value.ToString());
                writer.WriteElementString("Mode", rowdata[i, 1] = dgvRvRTestConditionData.Rows[i].Cells[1].Value.ToString());
                writer.WriteElementString("Channel", rowdata[i, 2] = dgvRvRTestConditionData.Rows[i].Cells[2].Value.ToString());
                writer.WriteElementString("SSID", rowdata[i, 3] = dgvRvRTestConditionData.Rows[i].Cells[3].Value.ToString());
                writer.WriteElementString("Start", rowdata[i, 4] = dgvRvRTestConditionData.Rows[i].Cells[4].Value.ToString());
                writer.WriteElementString("Stop", rowdata[i, 5] = dgvRvRTestConditionData.Rows[i].Cells[5].Value.ToString());
                writer.WriteElementString("Step", rowdata[i, 6] = dgvRvRTestConditionData.Rows[i].Cells[6].Value.ToString());
                writer.WriteElementString("TX_tst_File", rowdata[i, 6] = dgvRvRTestConditionData.Rows[i].Cells[7].Value.ToString());
                writer.WriteElementString("RX_tst_File", rowdata[i, 6] = dgvRvRTestConditionData.Rows[i].Cells[8].Value.ToString());
                writer.WriteElementString("BI_Dir_tst_File", rowdata[i, 6] = dgvRvRTestConditionData.Rows[i].Cells[9].Value.ToString());

                writer.WriteEndElement();
            }

            writer.WriteStartElement("Condition_Number");
            writer.WriteElementString("Number", dgvRvRTestConditionData.RowCount.ToString());
            writer.WriteEndElement();

            // End of write Function Test
            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public bool readXmlRvRTestCondition(string FileName)
        {
            int number = 0;

            XmlDocument doc = new XmlDocument();
            doc.Load(FileName);

            XmlNode nodeXml = doc.SelectSingleNode("CyberRouterATETC");
            if (nodeXml == null)
                return false;
            XmlElement element = (XmlElement)nodeXml;
            string strID = element.GetAttribute("Item");
            Debug.WriteLine(strID);
            if (strID.CompareTo("RvR Test Condition") != 0)
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
                    string Start = nodeTestCondition.SelectSingleNode("Start").InnerText;
                    string Stop = nodeTestCondition.SelectSingleNode("Stop").InnerText;
                    string Step = nodeTestCondition.SelectSingleNode("Step").InnerText;
                    string TX_tst_File = nodeTestCondition.SelectSingleNode("TX_tst_File").InnerText;
                    string RX_tst_File = nodeTestCondition.SelectSingleNode("RX_tst_File").InnerText;
                    string BI_Dir_tst_File = nodeTestCondition.SelectSingleNode("BI_Dir_tst_File").InnerText;

                    Debug.WriteLine("Band: " + Band);
                    Debug.WriteLine("Mode: " + Mode);
                    Debug.WriteLine("Channel: " + Channel);
                    Debug.WriteLine("SSID: " + SSID);
                    Debug.WriteLine("Start: " + Start);
                    Debug.WriteLine("Stop: " + Stop);
                    Debug.WriteLine("Step: " + Step);
                    Debug.WriteLine("TX_tst_File: " + TX_tst_File);
                    Debug.WriteLine("RX_tst_File: " + RX_tst_File);
                    Debug.WriteLine("BI_Dir_tst_File: " + BI_Dir_tst_File);

                    string[] data = new string[] { Band, Mode, Channel, SSID, Start, Stop, Step, TX_tst_File, RX_tst_File, BI_Dir_tst_File};
                    dgvRvRTestConditionData.Rows.Add(data);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("/CyberRouterATETC/TestCondition/Condition_Read " + ex);
                }
            }

            // End of Read Test Condition configuration settings

            return true;
        }
        
        private void DefaultAttenuation()
        {
            Attenuation_buttonValue_2_4G = new decimal[] { 1, 2, 4, 4, 10, 20, 30, 30 };
            Attenuation_buttonValue_5G = new decimal[] { 1, 2, 4, 4, 10, 20, 40, 40 };     
        }
    }
}
