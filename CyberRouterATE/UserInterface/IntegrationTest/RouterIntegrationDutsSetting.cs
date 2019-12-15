///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RouterIntegrationDutsSetting.cs
///  Update         : 2018-10-23
///  Description    : Main function
///  Modified       : 2018-10-23 Initial version  
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
using System.IO;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        /*============================================================================================*/
        /*==================== Global Paramter and Delegate Function  Declaration ====================*/
        /*============================================================================================*/
        #region        
        
        #endregion

        /*============================================================================================*/
        /*========================== Controller Event Function Area   ================================*/
        /*============================================================================================*/
        #region

        private void btnRouterIntegrationDutsSettingGuiScriptExcelFile_Click(object sender, EventArgs e)
        {
            string filename = @"RouterGuiScriptFile.xlsx";
            string sFilter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            string sInitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";

            txtRouterIntegrationDutsSettingGuiScripExcelFileName.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }    

        private void btnRouterIntegrationDutsSettingAddSetting_Click(object sender, EventArgs e)
        {
            if (!CheckParameterRouterIntegrationDutsSetting())
            {
                return;
            }

            RouterIntegrationDutsSettingAddSetting();
        }

        private void btnRouterIntegrationDutsSettingEditSetting_Click(object sender, EventArgs e)
        {
            /* James: To prevent one cell is to be inserted more time when Edit button clicked once again. */
            if (hasDeleteButton == false)
            {
                hasDeleteButton = true;
                btnRouterIntegrationDutsSettingEditSetting.Text = "Cancel";
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.Width = 66;
                btn.HeaderText = "Action";
                btn.Text = "Delete";
                btn.Name = "Action";
                btn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                btn.UseColumnTextForButtonValue = true;

                dgvRouterIntegrationDutsSettingData.Columns.Insert(0, btn);
            }
            else
            {
                btnRouterIntegrationDutsSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterIntegrationDutsSettingData.Columns.Remove("Action");
            }
        }

        private void btnRouterIntegrationDutsSettingSaveSetting_Click(object sender, EventArgs e)
        {
            if (btnRouterIntegrationDutsSettingEditSetting.Text == "Cancel")
            {
                btnRouterIntegrationDutsSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterIntegrationDutsSettingData.Columns.Remove("Action");
            }

            if (dgvRouterIntegrationDutsSettingData.RowCount <= 1)
            {
                MessageBox.Show("Add Condition First!!", "Warning");
                return;
            }

            string filename = "RouterIntegrationDutsSetting";

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
                writeXmlRouterIntegrationDutsSetting(saveFileDialog1.FileName);
            }
        }

        private void btnRouterIntegrationDutsSettingLoadSetting_Click(object sender, EventArgs e)
        {
            if (btnRouterIntegrationDutsSettingEditSetting.Text.ToLower() == "cancel")
            {
                btnRouterIntegrationDutsSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterIntegrationDutsSettingData.Columns.Remove("Action");
            }

            if (dgvRouterIntegrationDutsSettingData.RowCount > 1)
            {
                if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                    return;
                else
                {
                    DataTable dt = (DataTable)dgvRouterIntegrationDutsSettingData.DataSource;
                    dgvRouterIntegrationDutsSettingData.Rows.Clear();
                    dgvRouterIntegrationDutsSettingData.DataSource = dt;
                }
            }

            string filename = string.Empty;
            filename = "RouterIntegrationDutsSetting";

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = filename;
            openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";
            // Set filter for file extension and default file extension
            openFileDialog1.Filter = "XML file|*.xml";

            // If the file name is not an empty string open it for opening.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName != "")
            {
                readXmlRouterIntegrationDutsSetting(openFileDialog1.FileName);
            }

            //dgvRouterIntegrationDutsSettingData.AutoResizeColumns();
            //dgvRouterIntegrationDutsSettingData.ScrollBars = ScrollBars.Both;
        }

        private void btnRouterIntegrationDutsSettingMoveUp_Click(object sender, EventArgs e)
        {
            if (btnRouterIntegrationDutsSettingEditSetting.Text.ToLower() == "cancel")
            {
                btnRouterIntegrationDutsSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterIntegrationDutsSettingData.Columns.Remove("Action");
            }

            DataGridViewMoveUp(dgvRouterIntegrationDutsSettingData);
        }

        private void btnRouterIntegrationDutsSettingMoveDown_Click(object sender, EventArgs e)
        {
            if (btnRouterIntegrationDutsSettingEditSetting.Text.ToLower() == "cancel")
            {
                btnRouterIntegrationDutsSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterIntegrationDutsSettingData.Columns.Remove("Action");
            }

            DataGridViewMoveDown(dgvRouterIntegrationDutsSettingData);
        }

        private void dgvRouterIntegrationDutsSettingData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            /* James modified: Adding a condition to prevent exception error when remove null cell. */
            if (dgvRouterIntegrationDutsSettingData.Columns[e.ColumnIndex].Name == "Action" && hasDeleteButton == true && dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[0].Value != null)
            {
                dgvRouterIntegrationDutsSettingData.Rows.RemoveAt(e.RowIndex);
                if (e.ColumnIndex == 0 && (String)dgvRouterIntegrationDutsSettingData.Rows[0].Cells[0].Value == String.Empty)
                    dgvRouterIntegrationDutsSettingData.Columns.RemoveAt(0);
            }
        }

        private void dgvRouterIntegrationDutsSettingData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (btnRouterIntegrationDutsSettingEditSetting.Text.ToLower() != "cancel")
            {  
                nudRouterIntegrationDutsSettingIndex.Value = Convert.ToDecimal(dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[0].Value);
                txtRouterIntegrationDutsSettingModelName.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[1].Value.ToString();
                txtRouterIntegrationDutsSettingSerialNumber.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[2].Value.ToString();
                txtRouterIntegrationDutsSettingSwVersion.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[3].Value.ToString();
                txtRouterIntegrationDutsSettingHwVersion.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[4].Value.ToString();
                mtbRouterIntegrationDutsSettingIpAddress.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[5].Value.ToString();
                txtRouterIntegrationDutsSetting24gSsid.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[6].Value.ToString();
                txtRouterIntegrationDutsSetting5gSsid.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[7].Value.ToString();
                mtbRouterIntegrationDutsSettingPcIpAddress.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[8].Value.ToString();
                nudRouterIntegrationDutsSettingwitchPort.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[9].Value.ToString();
                cboxRouterIntegrationDutsSettingComPort.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[10].Value.ToString();
                mtbRouterIntegrationDutsSettingMacAddress.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[11].Value.ToString();
                txtRouterIntegrationDutsSettingGuiScripExcelFileName.Text = dgvRouterIntegrationDutsSettingData.Rows[e.RowIndex].Cells[12].Value.ToString();
            }
        }    
  
        #endregion

        /*============================================================================================*/
        /*=================================== Main Function Area =====================================*/
        /*============================================================================================*/
        #region Main Flow Function

        private void InitRouterIntegrationDutsSetting()
        {
            hasDeleteButton = false;
            SetupDataGridViewRouterIntegrationDutsSetting();

            /* Load the Data Set last time */
            string filename = System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterIntegrationDutsSetting.xml";

            if (File.Exists(filename))
                readXmlRouterIntegrationDutsSetting(filename);

        }

        public void SetupDataGridViewRouterIntegrationDutsSetting()
        {
            dgvRouterIntegrationDutsSettingData.ColumnCount = 13;
            dgvRouterIntegrationDutsSettingData.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvRouterIntegrationDutsSettingData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvRouterIntegrationDutsSettingData.Name = "Router Chamber Performance Duts Setting";
            //dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgvRouterIntegrationDutsSettingData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvRouterIntegrationDutsSettingData.Columns[0].Name = "Index";
            dgvRouterIntegrationDutsSettingData.Columns[1].Name = "Model Name";
            dgvRouterIntegrationDutsSettingData.Columns[2].Name = "Serial Number";
            dgvRouterIntegrationDutsSettingData.Columns[3].Name = "SW Version";
            dgvRouterIntegrationDutsSettingData.Columns[4].Name = "HW Version";
            dgvRouterIntegrationDutsSettingData.Columns[5].Name = "IP Address";
            dgvRouterIntegrationDutsSettingData.Columns[6].Name = "2.4G SSID";
            dgvRouterIntegrationDutsSettingData.Columns[7].Name = "5G SSID";
            dgvRouterIntegrationDutsSettingData.Columns[8].Name = "Control PC IP";
            dgvRouterIntegrationDutsSettingData.Columns[9].Name = "Switch Port";
            dgvRouterIntegrationDutsSettingData.Columns[10].Name = "ComPort";
            dgvRouterIntegrationDutsSettingData.Columns[11].Name = "MAC Address";
            dgvRouterIntegrationDutsSettingData.Columns[12].Name = "GUI Script File";            

            dgvRouterIntegrationDutsSettingData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvRouterIntegrationDutsSettingData.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvRouterIntegrationDutsSettingData.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False; //標題列換行, true -換行, false-不換行
            dgvRouterIntegrationDutsSettingData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//標題列置中

            //dgvRouterIntegrationDutsSettingData.Columns[0].Width = 120;
            //dgvRouterIntegrationDutsSettingData.Columns[1].Width = 120;
            //dgvRouterIntegrationDutsSettingData.Columns[2].Width = 120;
            //dgvRouterIntegrationDutsSettingData.Columns[3].Width = 80;
            //dgvRouterIntegrationDutsSettingData.Columns[4].Width = 80;
            //dgvRouterIntegrationDutsSettingData.Columns[5].Width = 80;            
            //dgvRouterIntegrationDutsSettingData.Columns[6].Width = 80;
            //dgvRouterIntegrationDutsSettingData.Columns[7].Width = 80;            
            //dgvRouterIntegrationDutsSettingData.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            //dgvRouterIntegrationDutsSettingData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            //dgvRouterIntegrationDutsSettingData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; //Use the column width setting
            //dgvRouterIntegrationDutsSettingData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            //FillAllColumns(dgvRouterIntegrationDutsSettingData, true);
        }

        #endregion

        #region Sub-Function

        private bool CheckParameterRouterIntegrationDutsSetting()
        {
            /* Check all data are not empty */
            if (nudRouterIntegrationDutsSettingIndex.Value.ToString() == "")
            {
                MessageBox.Show("Index can't be Empty!!!");
                return false;
            }

            //====== Model Info ======//
            if (txtRouterIntegrationDutsSettingModelName.Text == "")
            {
                MessageBox.Show("Duts Setting Serial Number can't be Empty!!!");                
                return false;
            }

            if (txtRouterIntegrationDutsSettingSerialNumber.Text == "")
            {
                MessageBox.Show("Duts Setting Serial Number can't be Empty!!!");
                return false;
            }

            if (txtRouterIntegrationDutsSettingSwVersion.Text == "")
            {
                MessageBox.Show("Duts Setting SW Version can't be Empty!!!");
                return false;
            }

            if (txtRouterIntegrationDutsSettingHwVersion.Text == "")
            {
                MessageBox.Show("Duts Setting HW Version can't be Empty!!!");
                return false;
            }

            if (mtbRouterIntegrationDutsSettingPcIpAddress.Text == "")
            {
                MessageBox.Show("Duts Setting Control PC IP Address can't be Empty!!!");
                return false;
            }

            //====== IP ======//
            if (mtbRouterIntegrationDutsSettingIpAddress.Text == "")
            {
                MessageBox.Show("Duts Setting IP Address can't be Empty!!!");
                return false;
            }

            if (mtbRouterIntegrationDutsSettingMacAddress.Text == "")
            {
                MessageBox.Show("Duts Setting MAC Address can't be Empty!!!");
                return false;
            }

            if (mtbRouterIntegrationDutsSettingPcIpAddress.Text == "")
            {
                MessageBox.Show("Duts Setting Control PC IP Address can't be Empty!!!");
                return false;
            }            

            //==== SSID =====//
            if (txtRouterIntegrationDutsSetting24gSsid.Text == "")
            {
                MessageBox.Show("Duts Setting 2.4G SSID can't be Empty!!!");
                return false;
            }

            if (txtRouterIntegrationDutsSetting5gSsid.Text == "")
            {
                MessageBox.Show("Duts Setting 5G SSID can't be Empty!!!");
                return false;
            }

            //==== Port =====//
            if (cboxRouterIntegrationDutsSettingComPort.Text == "")
            {
                MessageBox.Show("Duts Setting ComPort can't be Empty!!!");
                return false;
            }

            if (nudRouterIntegrationDutsSettingwitchPort.Text == "")
            {
                MessageBox.Show("Duts Setting Switch Port can't be Empty!!!");
                return false;
            }

            //==== GUI Script=====//
            if (txtRouterIntegrationDutsSettingGuiScripExcelFileName.Text == "")
            {
                MessageBox.Show("Duts Setting GUI Script Excel File Name can't be Empty!!!");
                return false;
            }

            /* Check all files exist */
            if (!File.Exists(txtRouterIntegrationDutsSettingGuiScripExcelFileName.Text))
            {
                MessageBox.Show("Duts Setting GUI Script Excel File Name doesn't exist!!!");
                return false;
            }            

            return true;
        }

        private void RouterIntegrationDutsSettingAddSetting()
        {
            /* Check if Delete button exist in the datagridview */
            if (hasDeleteButton)
            {
                hasDeleteButton = false;
                dgvRouterIntegrationDutsSettingData.Columns.RemoveAt(0);
                /* Reset the name of btn_TestCondition_Edit as Edit words. */
                btnRouterIntegrationDutsSettingEditSetting.Text = "Edit";
            }

            //dgvRouterIntegrationDutsSettingData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            //dgvRouterIntegrationDutsSettingData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            /* Add data to datagridview */
            // "Index", "Model Name", "Serial Number", "SW Version", "HW Version"
            // "IP Address", "SSID", "Control PC IP" , "Switch Port", "ComPort"
            // "MAC Address", "GUI Script File";
            string[] row = new string[]
            {            
                nudRouterIntegrationDutsSettingIndex.Value.ToString(),
                txtRouterIntegrationDutsSettingModelName.Text,
                txtRouterIntegrationDutsSettingSerialNumber.Text,
                txtRouterIntegrationDutsSettingSwVersion.Text,
                txtRouterIntegrationDutsSettingHwVersion.Text,
                mtbRouterIntegrationDutsSettingIpAddress.Text,
                txtRouterIntegrationDutsSetting24gSsid.Text,
                txtRouterIntegrationDutsSetting5gSsid.Text,
                mtbRouterIntegrationDutsSettingPcIpAddress.Text,
                nudRouterIntegrationDutsSettingwitchPort.Value.ToString(),
                cboxRouterIntegrationDutsSettingComPort.Text,
                mtbRouterIntegrationDutsSettingMacAddress.Text,
                txtRouterIntegrationDutsSettingGuiScripExcelFileName.Text,                
            };
            
            dgvRouterIntegrationDutsSettingData.Rows.Add(row);

            dgvRouterIntegrationDutsSettingData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvRouterIntegrationDutsSettingData.AutoResizeColumns();

            //if (!b_HasCheckBox)
            //{
            //    DataTable dt = (DataTable)dgvRouterIntegrationDutsSettingData.DataSource;
            //    dgvRouterIntegrationDutsSettingData.Rows.Clear();
            //    dgvRouterIntegrationDutsSettingData.DataSource = dt;


            //    DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
            //    dgvRouterIntegrationDutsSettingData.Columns.Insert(0, chk);
            //    b_HasCheckBox = true;
            //}

            //dgvRouterIntegrationDutsSettingData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
        }

        #endregion

        /*============================================================================================*/
        /*=================================== XML Function Area   ====================================*/
        /*============================================================================================*/
        #region        

        public void writeXmlRouterIntegrationDutsSetting(string FileName)
        {
            // ToDo: Needs to verify which test item is to be selected..
            int conditionCount = 0;
            string[,] rowdata = new string[dgvRouterIntegrationDutsSettingData.RowCount - 1, dgvRouterIntegrationDutsSettingData.ColumnCount];

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;

            XmlWriter writer = XmlWriter.Create(FileName, settings);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by the program.");
            writer.WriteStartElement("RouterATETC");
            writer.WriteAttributeString("Item", "Router Chamber Performance Duts Setting");
            
            ///
            /// Write Function Test settings
            /// 
            try
            {
                writer.WriteStartElement("DutsSetting");
                for (int i = 0; i < dgvRouterIntegrationDutsSettingData.RowCount - 1; i++)
                {
                    conditionCount++;
                    writer.WriteStartElement("Condition_" + (i + 1).ToString());

                    writer.WriteElementString("Index", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[0].Value.ToString());
                    writer.WriteElementString("ModelName", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[1].Value.ToString());
                    writer.WriteElementString("SerialNumber", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[2].Value.ToString());
                    writer.WriteElementString("SwVersion", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[3].Value.ToString());
                    writer.WriteElementString("HwVersion", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[4].Value.ToString());
                    writer.WriteElementString("IpAddress", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[5].Value.ToString());
                    writer.WriteElementString("SSID24g", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[6].Value.ToString());
                    writer.WriteElementString("SSID5g", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[7].Value.ToString());
                    writer.WriteElementString("ControlPcIp", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[8].Value.ToString());
                    writer.WriteElementString("SwitchPort", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[9].Value.ToString());
                    writer.WriteElementString("ComPort", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[10].Value.ToString());
                    writer.WriteElementString("MacAddress", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[11].Value.ToString());
                    writer.WriteElementString("GuiScriptFile", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[12].Value.ToString());
                    //writer.WriteElementString("TxTstFile", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[13].Value.ToString());
                    //writer.WriteElementString("RxTstFile", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[14].Value.ToString());
                    //writer.WriteElementString("BiTstFile", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[15].Value.ToString());
                    //writer.WriteElementString("FwFileName", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[12].Value.ToString());
                    //writer.WriteElementString("", dgvRouterIntegrationDutsSettingData.Rows[i].Cells[13].Value.ToString());

                    writer.WriteEndElement();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Write Test Condition: " + ex.ToString());
            }

            writer.WriteStartElement("Condition_Number");
            writer.WriteElementString("Number", dgvRouterIntegrationDutsSettingData.RowCount.ToString());
            writer.WriteEndElement();

            // End of write Function Test
            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public bool readXmlRouterIntegrationDutsSetting(string FileName)
        {
            int number = 0;

            //dgvRouterIntegrationDutsSettingData.Columns[dgvRouterIntegrationDutsSettingData.ColumnCount].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            XmlDocument doc = new XmlDocument();
            doc.Load(FileName);

            XmlNode nodeXml = doc.SelectSingleNode("RouterATETC");
            if (nodeXml == null)
                return false;
            XmlElement element = (XmlElement)nodeXml;
            string strID = element.GetAttribute("Item");
            Debug.WriteLine(strID);
            if (strID.CompareTo("Router Chamber Performance Duts Setting") != 0)
            {
                MessageBox.Show("This XML file is incorrect.", "Error");
                return false;
            }

            ///
            /// Read Function Test configuration settings
            ///

            XmlNode nodeTestConditionModel = doc.SelectSingleNode("/RouterATETC/DutsSetting/Condition_Number");
            try
            {
                string Number = nodeTestConditionModel.SelectSingleNode("Number").InnerText;
                Debug.WriteLine("Number: " + Name);
                number = Int32.Parse(Number);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATETC/DutsSetting//Condition_Number" + ex);
            }

            for (int i = 0; i < number; i++)
            {
                XmlNode nodeTestCondition = doc.SelectSingleNode("/RouterATETC/DutsSetting//Condition_" + (i + 1).ToString());

                try
                {
                    string Index = nodeTestCondition.SelectSingleNode("Index").InnerText;
                    string ModelName = nodeTestCondition.SelectSingleNode("ModelName").InnerText;
                    string SerialNumber = nodeTestCondition.SelectSingleNode("SerialNumber").InnerText;
                    string SwVersion = nodeTestCondition.SelectSingleNode("SwVersion").InnerText;
                    string HwVersion = nodeTestCondition.SelectSingleNode("HwVersion").InnerText;
                    string IpAddress = nodeTestCondition.SelectSingleNode("IpAddress").InnerText;
                    string SSID24g = nodeTestCondition.SelectSingleNode("SSID24g").InnerText;
                    string SSID5g = nodeTestCondition.SelectSingleNode("SSID5g").InnerText;
                    string ControlPcIp = nodeTestCondition.SelectSingleNode("ControlPcIp").InnerText;
                    string SwitchPort = nodeTestCondition.SelectSingleNode("SwitchPort").InnerText;
                    string ComPort = nodeTestCondition.SelectSingleNode("ComPort").InnerText;
                    string MacAddress = nodeTestCondition.SelectSingleNode("MacAddress").InnerText;
                    string GuiScriptFile = nodeTestCondition.SelectSingleNode("GuiScriptFile").InnerText;
                    //string TxTstFile = nodeTestCondition.SelectSingleNode("TxTstFile").InnerText;
                    //string RxTstFile = nodeTestCondition.SelectSingleNode("RxTstFile").InnerText;
                    //string BiTstFile = nodeTestCondition.SelectSingleNode("BiTstFile").InnerText;

                    Debug.WriteLine("Index: " + Index);
                    Debug.WriteLine("ModelName: " + ModelName);
                    Debug.WriteLine("SerialNumber: " + SerialNumber);
                    Debug.WriteLine("SwVersion: " + SwVersion);
                    Debug.WriteLine("HwVersion: " + HwVersion);
                    Debug.WriteLine("IpAddress: " + IpAddress);
                    Debug.WriteLine("SSID24g: " + SSID24g);
                    Debug.WriteLine("SSID5g: " + SSID5g);
                    Debug.WriteLine("ControlPcIp: " + ControlPcIp);
                    Debug.WriteLine("SwitchPort: " + SwitchPort);
                    Debug.WriteLine("ComPort: " + ComPort);
                    Debug.WriteLine("MacAddress: " + MacAddress);
                    Debug.WriteLine("GuiScriptFile: " + GuiScriptFile);
                    //Debug.WriteLine("TxTstFile: " + TxTstFile);
                    //Debug.WriteLine("RxTstFile: " + RxTstFile);
                    //Debug.WriteLine("BiTstFile: " + BiTstFile);

                    string[] data = new string[] { Index, ModelName, SerialNumber, SwVersion, HwVersion, IpAddress, SSID24g, SSID5g, ControlPcIp, SwitchPort, ComPort, MacAddress, GuiScriptFile };
                    dgvRouterIntegrationDutsSettingData.Rows.Add(data);

                }
                catch (Exception ex)
                {
                    Debug.WriteLine("/RouterATETC/DutsSetting//Condition_Read " + ex);
                }
            }

            // End of Read Test Condition configuration settings
            //FillAllColumns(dgvRouterIntegrationDutsSettingData, true);
            return true;
        }

        #endregion
        
        /*============================================================================================*/
        /*======================================= The End   ==========================================*/
        /*============================================================================================*/


    }
}
