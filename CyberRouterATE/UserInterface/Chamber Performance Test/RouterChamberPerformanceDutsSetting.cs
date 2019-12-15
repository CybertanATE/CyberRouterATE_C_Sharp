///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RouterChamberPerformanceDutsSetting.cs
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

        int i_RouterChamberPerformanceDutsSettingIndex = 1;
        bool b_HasCheckBox = false;

        #endregion

        /*============================================================================================*/
        /*========================== Controller Event Function Area   ================================*/
        /*============================================================================================*/
        #region

        private void btnRouterChamberPerformanceDutsSettingGuiScriptExcelFile_Click(object sender, EventArgs e)
        {
            string filename = @"RouterGuiScriptFile.xlsx";
            string sFilter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            string sInitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";

            txtRouterChamberPerformanceDutsSettingGuiScripExcelFileName.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }    

        private void btnRouterChamberPerformanceDutsSettingAddSetting_Click(object sender, EventArgs e)
        {
            if (!CheckParameterRouterChamberPerformanceDutsSetting())
            {
                return;
            }

            RouterChamberPerformanceDutsSettingAddSetting();
        }

        private void btnRouterChamberPerformanceDutsSettingEditSetting_Click(object sender, EventArgs e)
        {
            /* James: To prevent one cell is to be inserted more time when Edit button clicked once again. */
            if (hasDeleteButton == false)
            {
                hasDeleteButton = true;
                btnRouterChamberPerformanceDutsSettingEditSetting.Text = "Cancel";
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.Width = 66;
                btn.HeaderText = "Action";
                btn.Text = "Delete";
                btn.Name = "Action";
                btn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                btn.UseColumnTextForButtonValue = true;

                dgvRouterChamberPerformanceDutsSettingData.Columns.Insert(0, btn);
            }
            else
            {
                btnRouterChamberPerformanceDutsSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterChamberPerformanceDutsSettingData.Columns.Remove("Action");
            }
        }

        private void btnRouterChamberPerformanceDutsSettingSaveSetting_Click(object sender, EventArgs e)
        {
            if (btnRouterChamberPerformanceDutsSettingEditSetting.Text == "Cancel")
            {
                btnRouterChamberPerformanceDutsSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterChamberPerformanceDutsSettingData.Columns.Remove("Action");
            }

            if (dgvRouterChamberPerformanceDutsSettingData.RowCount <= 1)
            {
                MessageBox.Show("Add Condition First!!", "Warning");
                return;
            }

            string filename = "RouterChamberPerformanceDutsSetting";

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
                writeXmlRouterChamberPerformanceDutsSetting(saveFileDialog1.FileName);
            }
        }

        private void btnRouterChamberPerformanceDutsSettingLoadSetting_Click(object sender, EventArgs e)
        {
            if (btnRouterChamberPerformanceDutsSettingEditSetting.Text.ToLower() == "cancel")
            {
                btnRouterChamberPerformanceDutsSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterChamberPerformanceDutsSettingData.Columns.Remove("Action");
            }

            if (dgvRouterChamberPerformanceDutsSettingData.RowCount > 1)
            {
                if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                    return;
                else
                {
                    DataTable dt = (DataTable)dgvRouterChamberPerformanceDutsSettingData.DataSource;
                    dgvRouterChamberPerformanceDutsSettingData.Rows.Clear();
                    dgvRouterChamberPerformanceDutsSettingData.DataSource = dt;
                }
            }

            string filename = string.Empty;
            filename = "RouterChamberPerformanceDutsSetting";

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = filename;
            openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";
            // Set filter for file extension and default file extension
            openFileDialog1.Filter = "XML file|*.xml";

            // If the file name is not an empty string open it for opening.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName != "")
            {
                readXmlRouterChamberPerformanceDutsSetting(openFileDialog1.FileName);
            }

            //dgvRouterChamberPerformanceDutsSettingData.AutoResizeColumns();
            //dgvRouterChamberPerformanceDutsSettingData.ScrollBars = ScrollBars.Both;
        }

        private void btnRouterChamberPerformanceDutsSettingMoveUp_Click(object sender, EventArgs e)
        {
            if (btnRouterChamberPerformanceDutsSettingEditSetting.Text.ToLower() == "cancel")
            {
                btnRouterChamberPerformanceDutsSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterChamberPerformanceDutsSettingData.Columns.Remove("Action");
            }

            DataGridViewMoveUp(dgvRouterChamberPerformanceDutsSettingData);
        }

        private void btnRouterChamberPerformanceDutsSettingMoveDown_Click(object sender, EventArgs e)
        {
            if (btnRouterChamberPerformanceDutsSettingEditSetting.Text.ToLower() == "cancel")
            {
                btnRouterChamberPerformanceDutsSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterChamberPerformanceDutsSettingData.Columns.Remove("Action");
            }

            DataGridViewMoveDown(dgvRouterChamberPerformanceDutsSettingData);
        }

        private void dgvRouterChamberPerformanceDutsSettingData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            /* James modified: Adding a condition to prevent exception error when remove null cell. */
            if (dgvRouterChamberPerformanceDutsSettingData.Columns[e.ColumnIndex].Name == "Action" && hasDeleteButton == true && dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[0].Value != null)
            {
                dgvRouterChamberPerformanceDutsSettingData.Rows.RemoveAt(e.RowIndex);
                if (e.ColumnIndex == 0 && (String)dgvRouterChamberPerformanceDutsSettingData.Rows[0].Cells[0].Value == String.Empty)
                    dgvRouterChamberPerformanceDutsSettingData.Columns.RemoveAt(0);
            }
        }

        private void dgvRouterChamberPerformanceDutsSettingData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (btnRouterChamberPerformanceDutsSettingEditSetting.Text.ToLower() != "cancel")
            {  
                nudRouterChamberPerformanceDutsSettingIndex.Value = Convert.ToDecimal(dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[0].Value);
                txtRouterChamberPerformanceDutsSettingModelName.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[1].Value.ToString();
                txtRouterChamberPerformanceDutsSettingSerialNumber.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[2].Value.ToString();
                txtRouterChamberPerformanceDutsSettingSwVersion.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[3].Value.ToString();
                txtRouterChamberPerformanceDutsSettingHwVersion.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[4].Value.ToString();
                mtbRouterChamberPerformanceDutsSettingIpAddress.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[5].Value.ToString();
                txtRouterChamberPerformanceDutsSetting24gSsid.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[6].Value.ToString();
                txtRouterChamberPerformanceDutsSetting5gSsid.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[7].Value.ToString();
                mtbRouterChamberPerformanceDutsSettingPcIpAddress.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[8].Value.ToString();
                nudRouterChamberPerformanceDutsSettingwitchPort.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[9].Value.ToString();
                cboxRouterChamberPerformanceDutsSettingComPort.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[10].Value.ToString();
                mtbRouterChamberPerformanceDutsSettingMacAddress.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[11].Value.ToString();
                txtRouterChamberPerformanceDutsSettingGuiScripExcelFileName.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[12].Value.ToString();
                txtRouterChamberPerformanceDutsSettingTxTstFile.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[13].Value.ToString();
                txtRouterChamberPerformanceDutsSettingRxTstFile.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[14].Value.ToString();
                txtRouterChamberPerformanceDutsSettingBiTstFile.Text = dgvRouterChamberPerformanceDutsSettingData.Rows[e.RowIndex].Cells[15].Value.ToString();
            }
        }

        private void btnRouterChamberPerformanceDutsSettingTxTstFile_Click(object sender, EventArgs e)
        {
            string filename = @"Tx.tst";
            string sFilter = "Tst files (*.tst)|*.tst|All files (*.*)|*.*";
            string sInitialDirectory = System.Windows.Forms.Application.StartupPath + @"\importData\";

            txtRouterChamberPerformanceDutsSettingTxTstFile.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }    

        private void btnRouterChamberPerformanceDutsSettingRxTstFile_Click(object sender, EventArgs e)
        {
            string filename = @"Rx.tst";
            string sFilter = "Tst files (*.tst)|*.tst|All files (*.*)|*.*";
            string sInitialDirectory = System.Windows.Forms.Application.StartupPath + @"\importData\";

            txtRouterChamberPerformanceDutsSettingRxTstFile.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }    

        private void btnRouterChamberPerformanceDutsSettingBiTstFile_Click(object sender, EventArgs e)
        {
            string filename = @"Bi.tst";
            string sFilter = "Tst files (*.tst)|*.tst|All files (*.*)|*.*";
            string sInitialDirectory = System.Windows.Forms.Application.StartupPath + @"\importData\";

            txtRouterChamberPerformanceDutsSettingBiTstFile.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }    
        
        
        #endregion

        /*============================================================================================*/
        /*=================================== Main Function Area =====================================*/
        /*============================================================================================*/
        #region Main Flow Function

        private void InitRouterChamberPerformanceDutsSetting()
        {
            hasDeleteButton = false;
            SetupDataGridViewRouterChamberPerformanceDutsSetting();
        }

        public void SetupDataGridViewRouterChamberPerformanceDutsSetting()
        {
            dgvRouterChamberPerformanceDutsSettingData.ColumnCount = 16;
            dgvRouterChamberPerformanceDutsSettingData.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvRouterChamberPerformanceDutsSettingData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvRouterChamberPerformanceDutsSettingData.Name = "Router Chamber Performance Duts Setting";
            //dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgvRouterChamberPerformanceDutsSettingData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvRouterChamberPerformanceDutsSettingData.Columns[0].Name = "Index";
            dgvRouterChamberPerformanceDutsSettingData.Columns[1].Name = "Model Name";
            dgvRouterChamberPerformanceDutsSettingData.Columns[2].Name = "Serial Number";
            dgvRouterChamberPerformanceDutsSettingData.Columns[3].Name = "SW Version";
            dgvRouterChamberPerformanceDutsSettingData.Columns[4].Name = "HW Version";
            dgvRouterChamberPerformanceDutsSettingData.Columns[5].Name = "IP Address";
            dgvRouterChamberPerformanceDutsSettingData.Columns[6].Name = "2.4G SSID";
            dgvRouterChamberPerformanceDutsSettingData.Columns[7].Name = "5G SSID";
            dgvRouterChamberPerformanceDutsSettingData.Columns[8].Name = "Control PC IP";
            dgvRouterChamberPerformanceDutsSettingData.Columns[9].Name = "Switch Port";
            dgvRouterChamberPerformanceDutsSettingData.Columns[10].Name = "ComPort";
            dgvRouterChamberPerformanceDutsSettingData.Columns[11].Name = "MAC Address";
            dgvRouterChamberPerformanceDutsSettingData.Columns[12].Name = "GUI Script File";
            dgvRouterChamberPerformanceDutsSettingData.Columns[13].Name = "TX tst File";
            dgvRouterChamberPerformanceDutsSettingData.Columns[14].Name = "RX tst File";
            dgvRouterChamberPerformanceDutsSettingData.Columns[15].Name = "Bi tst File";

            dgvRouterChamberPerformanceDutsSettingData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvRouterChamberPerformanceDutsSettingData.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvRouterChamberPerformanceDutsSettingData.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False; //標題列換行, true -換行, false-不換行
            dgvRouterChamberPerformanceDutsSettingData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//標題列置中

            //dgvRouterChamberPerformanceDutsSettingData.Columns[0].Width = 120;
            //dgvRouterChamberPerformanceDutsSettingData.Columns[1].Width = 120;
            //dgvRouterChamberPerformanceDutsSettingData.Columns[2].Width = 120;
            //dgvRouterChamberPerformanceDutsSettingData.Columns[3].Width = 80;
            //dgvRouterChamberPerformanceDutsSettingData.Columns[4].Width = 80;
            //dgvRouterChamberPerformanceDutsSettingData.Columns[5].Width = 80;            
            //dgvRouterChamberPerformanceDutsSettingData.Columns[6].Width = 80;
            //dgvRouterChamberPerformanceDutsSettingData.Columns[7].Width = 80;            
            //dgvRouterChamberPerformanceDutsSettingData.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            //dgvRouterChamberPerformanceDutsSettingData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            //dgvRouterChamberPerformanceDutsSettingData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; //Use the column width setting
            //dgvRouterChamberPerformanceDutsSettingData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            //FillAllColumns(dgvRouterChamberPerformanceDutsSettingData, true);
        }

        #endregion

        #region Sub-Function

        private bool CheckParameterRouterChamberPerformanceDutsSetting()
        {
            /* Check all data are not empty */
            if (nudRouterChamberPerformanceDutsSettingIndex.Value.ToString() == "")
            {
                MessageBox.Show("Index can't be Empty!!!");
                return false;
            }

            //====== Model Info ======//
            if (txtRouterChamberPerformanceDutsSettingModelName.Text == "")
            {
                MessageBox.Show("Serial Number can't be Empty!!!");
                return false;
            }

            if (txtRouterChamberPerformanceDutsSettingSerialNumber.Text == "")
            {
                MessageBox.Show("Serial Number can't be Empty!!!");
                return false;
            }

            if (txtRouterChamberPerformanceDutsSettingSwVersion.Text == "")
            {
                MessageBox.Show("SW Version can't be Empty!!!");
                return false;
            }

            if (txtRouterChamberPerformanceDutsSettingHwVersion.Text == "")
            {
                MessageBox.Show("HW Version can't be Empty!!!");
                return false;
            }

            if (mtbRouterChamberPerformanceDutsSettingPcIpAddress.Text == "")
            {
                MessageBox.Show("Control PC IP Address can't be Empty!!!");
                return false;
            }

            //====== IP ======//
            if (mtbRouterChamberPerformanceDutsSettingIpAddress.Text == "")
            {
                MessageBox.Show("IP Address can't be Empty!!!");
                return false;
            }

            if (mtbRouterChamberPerformanceDutsSettingMacAddress.Text == "")
            {
                MessageBox.Show("MAC Address can't be Empty!!!");
                return false;
            }

            if (mtbRouterChamberPerformanceDutsSettingPcIpAddress.Text == "")
            {
                MessageBox.Show("Control PC IP Address can't be Empty!!!");
                return false;
            }

            if (txtRouterChamberPerformanceDutsSettingTxTstFile.Text == "")
            {
                MessageBox.Show("TX tst File can't be Empty!!!");
                return false;
            }

            if (txtRouterChamberPerformanceDutsSettingRxTstFile.Text == "")
            {
                MessageBox.Show("RX tst File can't be Empty!!!");
                return false;
            }

            if (txtRouterChamberPerformanceDutsSettingBiTstFile.Text == "")
            {
                MessageBox.Show("Bi-Direction tst File can't be Empty!!!");
                return false;
            }

            //==== SSID =====//
            if (txtRouterChamberPerformanceDutsSetting24gSsid.Text == "")
            {
                MessageBox.Show("2.4G SSID can't be Empty!!!");
                return false;
            }

            if (txtRouterChamberPerformanceDutsSetting5gSsid.Text == "")
            {
                MessageBox.Show("5G SSID can't be Empty!!!");
                return false;
            }

            //==== Port =====//
            if (cboxRouterChamberPerformanceDutsSettingComPort.Text == "")
            {
                MessageBox.Show("ComPort can't be Empty!!!");
                return false;
            }

            if (nudRouterChamberPerformanceDutsSettingwitchPort.Text == "")
            {
                MessageBox.Show("Switch Port can't be Empty!!!");
                return false;
            }

            //==== GUI Script=====//
            if (txtRouterChamberPerformanceDutsSettingGuiScripExcelFileName.Text == "")
            {
                MessageBox.Show("GUI Script Excel File Name can't be Empty!!!");
                return false;
            }

            /* Check all files exist */
            if (!File.Exists(txtRouterChamberPerformanceDutsSettingGuiScripExcelFileName.Text))
            {
                MessageBox.Show("GUI Script Excel File Name doesn't exist!!!");
                return false;
            }

            if (!File.Exists(txtRouterChamberPerformanceDutsSettingTxTstFile.Text))
            {
                MessageBox.Show("TX tst File doesn't exist!!!");
                return false;
            }

            if (!File.Exists(txtRouterChamberPerformanceDutsSettingRxTstFile.Text))
            {
                MessageBox.Show("RX tst File doesn't exist!!!");
                return false;
            }

            if (!File.Exists(txtRouterChamberPerformanceDutsSettingBiTstFile.Text))
            {
                MessageBox.Show("Bi-Direction tst File doesn't exist!!!");
                return false;
            }

            return true;
        }

        private void RouterChamberPerformanceDutsSettingAddSetting()
        {
            /* Check if Delete button exist in the datagridview */
            if (hasDeleteButton)
            {
                hasDeleteButton = false;
                dgvRouterChamberPerformanceDutsSettingData.Columns.RemoveAt(0);
                /* Reset the name of btn_TestCondition_Edit as Edit words. */
                btnRouterChamberPerformanceDutsSettingEditSetting.Text = "Edit";
            }

            //dgvRouterChamberPerformanceDutsSettingData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            //dgvRouterChamberPerformanceDutsSettingData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            /* Add data to datagridview */
            // "Index", "Model Name", "Serial Number", "SW Version", "HW Version"
            // "IP Address", "SSID", "Control PC IP" , "Switch Port", "ComPort"
            // "MAC Address", "GUI Script File";
            string[] row = new string[]
            {            
                nudRouterChamberPerformanceDutsSettingIndex.Value.ToString(),
                txtRouterChamberPerformanceDutsSettingModelName.Text,
                txtRouterChamberPerformanceDutsSettingSerialNumber.Text,
                txtRouterChamberPerformanceDutsSettingSwVersion.Text,
                txtRouterChamberPerformanceDutsSettingHwVersion.Text,
                mtbRouterChamberPerformanceDutsSettingIpAddress.Text,
                txtRouterChamberPerformanceDutsSetting24gSsid.Text,
                txtRouterChamberPerformanceDutsSetting5gSsid.Text,
                mtbRouterChamberPerformanceDutsSettingPcIpAddress.Text,
                nudRouterChamberPerformanceDutsSettingwitchPort.Value.ToString(),
                cboxRouterChamberPerformanceDutsSettingComPort.Text,
                mtbRouterChamberPerformanceDutsSettingMacAddress.Text,
                txtRouterChamberPerformanceDutsSettingGuiScripExcelFileName.Text,
                txtRouterChamberPerformanceDutsSettingTxTstFile.Text,
                txtRouterChamberPerformanceDutsSettingRxTstFile.Text,
                txtRouterChamberPerformanceDutsSettingBiTstFile.Text
            };

            i_RouterChamberPerformanceDutsSettingIndex++; //index, auto-increment by 1
            dgvRouterChamberPerformanceDutsSettingData.Rows.Add(row);

            dgvRouterChamberPerformanceDutsSettingData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvRouterChamberPerformanceDutsSettingData.AutoResizeColumns();

            //if (!b_HasCheckBox)
            //{
            //    DataTable dt = (DataTable)dgvRouterChamberPerformanceDutsSettingData.DataSource;
            //    dgvRouterChamberPerformanceDutsSettingData.Rows.Clear();
            //    dgvRouterChamberPerformanceDutsSettingData.DataSource = dt;


            //    DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
            //    dgvRouterChamberPerformanceDutsSettingData.Columns.Insert(0, chk);
            //    b_HasCheckBox = true;
            //}

            //dgvRouterChamberPerformanceDutsSettingData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
        }

        #endregion

        /*============================================================================================*/
        /*=================================== XML Function Area   ====================================*/
        /*============================================================================================*/
        #region        

        public void writeXmlRouterChamberPerformanceDutsSetting(string FileName)
        {
            // ToDo: Needs to verify which test item is to be selected..
            int conditionCount = 0;
            string[,] rowdata = new string[dgvRouterChamberPerformanceDutsSettingData.RowCount - 1, dgvRouterChamberPerformanceDutsSettingData.ColumnCount];

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
                for (int i = 0; i < dgvRouterChamberPerformanceDutsSettingData.RowCount - 1; i++)
                {
                    conditionCount++;
                    writer.WriteStartElement("Condition_" + (i + 1).ToString());

                    writer.WriteElementString("Index", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[0].Value.ToString());
                    writer.WriteElementString("ModelName", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[1].Value.ToString());
                    writer.WriteElementString("SerialNumber", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[2].Value.ToString());
                    writer.WriteElementString("SwVersion", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[3].Value.ToString());
                    writer.WriteElementString("HwVersion", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[4].Value.ToString());
                    writer.WriteElementString("IpAddress", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[5].Value.ToString());
                    writer.WriteElementString("SSID24g", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[6].Value.ToString());
                    writer.WriteElementString("SSID5g", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[7].Value.ToString());
                    writer.WriteElementString("ControlPcIp", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[8].Value.ToString());
                    writer.WriteElementString("SwitchPort", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[9].Value.ToString());
                    writer.WriteElementString("ComPort", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[10].Value.ToString());
                    writer.WriteElementString("MacAddress", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[11].Value.ToString());
                    writer.WriteElementString("GuiScriptFile", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[12].Value.ToString());
                    writer.WriteElementString("TxTstFile", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[13].Value.ToString());
                    writer.WriteElementString("RxTstFile", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[14].Value.ToString());
                    writer.WriteElementString("BiTstFile", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[15].Value.ToString());
                    //writer.WriteElementString("FwFileName", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[12].Value.ToString());
                    //writer.WriteElementString("", dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[13].Value.ToString());

                    writer.WriteEndElement();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Write Test Condition: " + ex.ToString());
            }

            writer.WriteStartElement("Condition_Number");
            writer.WriteElementString("Number", dgvRouterChamberPerformanceDutsSettingData.RowCount.ToString());
            writer.WriteEndElement();

            // End of write Function Test
            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public bool readXmlRouterChamberPerformanceDutsSetting(string FileName)
        {
            int number = 0;

            //dgvRouterChamberPerformanceDutsSettingData.Columns[dgvRouterChamberPerformanceDutsSettingData.ColumnCount].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

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
                    string TxTstFile = nodeTestCondition.SelectSingleNode("TxTstFile").InnerText;
                    string RxTstFile = nodeTestCondition.SelectSingleNode("RxTstFile").InnerText;
                    string BiTstFile = nodeTestCondition.SelectSingleNode("BiTstFile").InnerText;

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
                    Debug.WriteLine("TxTstFile: " + TxTstFile);
                    Debug.WriteLine("RxTstFile: " + RxTstFile);
                    Debug.WriteLine("BiTstFile: " + BiTstFile);

                    string[] data = new string[] { Index, ModelName, SerialNumber, SwVersion, HwVersion, IpAddress, SSID24g, SSID5g, ControlPcIp, SwitchPort, ComPort, MacAddress, GuiScriptFile, TxTstFile, RxTstFile, BiTstFile };
                    dgvRouterChamberPerformanceDutsSettingData.Rows.Add(data);

                }
                catch (Exception ex)
                {
                    Debug.WriteLine("/RouterATETC/DutsSetting//Condition_Read " + ex);
                }
            }

            // End of Read Test Condition configuration settings
            //FillAllColumns(dgvRouterChamberPerformanceDutsSettingData, true);
            return true;
        }

        #endregion
        
        /*============================================================================================*/
        /*======================================= The End   ==========================================*/
        /*============================================================================================*/


    }
}
