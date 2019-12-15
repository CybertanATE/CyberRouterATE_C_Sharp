///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RouterIntegrationTestCaseSetting.cs
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

        private void btnRouterIntegrationTestCaseSettingFinalExcelFile_Click(object sender, EventArgs e)
        {
            string filename = @"RouterTestCaseExcelFile.xlsx";
            string sFilter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            string sInitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";

            txtRouterIntegrationTestCaseSettingFinalExcelFile.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }

        private void btnRouterIntegrationTestCaseSettingStepExcelFile_Click(object sender, EventArgs e)
        {
            string filename = @"RouterTestCaseExcelFile.xlsx";
            string sFilter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            string sInitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";

            txtRouterIntegrationTestCaseSettingStepExcelFile.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }

        private void btnRouterIntegrationTestCaseSettingAddSetting_Click(object sender, EventArgs e)
        {
            if (!CheckParameterRouterIntegrationTestCaseSetting())
            {
                return;
            }

            RouterIntegrationTestCaseSettingAddSetting();
        }

        private void btnRouterIntegrationTestCaseSettingEditSetting_Click(object sender, EventArgs e)
        {
            /* James: To prevent one cell is to be inserted more time when Edit button clicked once again. */
            if (hasDeleteButton == false)
            {
                hasDeleteButton = true;
                btnRouterIntegrationTestCaseSettingEditSetting.Text = "Cancel";
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.Width = 66;
                btn.HeaderText = "Action";
                btn.Text = "Delete";
                btn.Name = "Action";
                btn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                btn.UseColumnTextForButtonValue = true;

                dgvRouterIntegrationTestCaseSettingData.Columns.Insert(0, btn);
            }
            else
            {
                btnRouterIntegrationTestCaseSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterIntegrationTestCaseSettingData.Columns.Remove("Action");
            }
        }

        private void btnRouterIntegrationTestCaseSettingSaveSetting_Click(object sender, EventArgs e)
        {
            if (btnRouterIntegrationTestCaseSettingEditSetting.Text == "Cancel")
            {
                btnRouterIntegrationTestCaseSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterIntegrationTestCaseSettingData.Columns.Remove("Action");
            }

            if (dgvRouterIntegrationTestCaseSettingData.RowCount <= 1)
            {
                MessageBox.Show("Add Condition First!!", "Warning");
                return;
            }

            string filename = "RouterIntegrationTestCaseSetting";

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
                writeXmlRouterIntegrationTestCaseSetting(saveFileDialog1.FileName);
            }
        }

        private void btnRouterIntegrationTestCaseSettingLoadSetting_Click(object sender, EventArgs e)
        {
            if (btnRouterIntegrationTestCaseSettingEditSetting.Text.ToLower() == "cancel")
            {
                btnRouterIntegrationTestCaseSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterIntegrationTestCaseSettingData.Columns.Remove("Action");
            }

            if (dgvRouterIntegrationTestCaseSettingData.RowCount > 1)
            {
                if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                    return;
                else
                {
                    DataTable dt = (DataTable)dgvRouterIntegrationTestCaseSettingData.DataSource;
                    dgvRouterIntegrationTestCaseSettingData.Rows.Clear();
                    dgvRouterIntegrationTestCaseSettingData.DataSource = dt;
                }
            }

            string filename = string.Empty;
            filename = "RouterIntegrationTestCaseSetting";

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = filename;
            openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";
            // Set filter for file extension and default file extension
            openFileDialog1.Filter = "XML file|*.xml";

            // If the file name is not an empty string open it for opening.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName != "")
            {
                readXmlRouterIntegrationTestCaseSetting(openFileDialog1.FileName);
            }

            //dgvRouterIntegrationTestCaseSettingData.AutoResizeColumns();
            //dgvRouterIntegrationTestCaseSettingData.ScrollBars = ScrollBars.Both;
        }

        private void btnRouterIntegrationTestCaseSettingMoveUp_Click(object sender, EventArgs e)
        {
            if (btnRouterIntegrationTestCaseSettingEditSetting.Text.ToLower() == "cancel")
            {
                btnRouterIntegrationTestCaseSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterIntegrationTestCaseSettingData.Columns.Remove("Action");
            }

            DataGridViewMoveUp(dgvRouterIntegrationTestCaseSettingData);
        }

        private void btnRouterIntegrationTestCaseSettingMoveDown_Click(object sender, EventArgs e)
        {
            if (btnRouterIntegrationTestCaseSettingEditSetting.Text.ToLower() == "cancel")
            {
                btnRouterIntegrationTestCaseSettingEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvRouterIntegrationTestCaseSettingData.Columns.Remove("Action");
            }

            DataGridViewMoveDown(dgvRouterIntegrationTestCaseSettingData);
        }

        private void dgvRouterIntegrationTestCaseSettingData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            /* James modified: Adding a condition to prevent exception error when remove null cell. */
            if (dgvRouterIntegrationTestCaseSettingData.Columns[e.ColumnIndex].Name == "Action" && hasDeleteButton == true && dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[0].Value != null)
            {
                dgvRouterIntegrationTestCaseSettingData.Rows.RemoveAt(e.RowIndex);
                if (e.ColumnIndex == 0 && (String)dgvRouterIntegrationTestCaseSettingData.Rows[0].Cells[0].Value == String.Empty)
                    dgvRouterIntegrationTestCaseSettingData.Columns.RemoveAt(0);
            }
        }

        private void dgvRouterIntegrationTestCaseSettingData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (btnRouterIntegrationTestCaseSettingEditSetting.Text.ToLower() != "cancel")
            {  
                nudRouterIntegrationTestCaseSettingIndex.Value = Convert.ToDecimal(dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[0].Value);
                txtRouterIntegrationTestCaseSettingItemName.Text = dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[1].Value.ToString();
                chkRouterIntegrationTestCaseSettingDutResetToDefaultAfterTest.Checked = (dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[2].Value.ToString() == "Y" ? true : false);
                chkRouterIntegrationTestCaseSettingDutRebootAfterTest.Checked = (dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[3].Value.ToString() == "Y"? true: false);
                nudRouterIntegrationTestCaseSettingWaitTimeAfterTest.Value = Convert.ToDecimal(dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[4].Value);
                txtRouterIntegrationTestCaseSettingFinalExcelFile.Text = dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[5].Value.ToString();
                txtRouterIntegrationTestCaseSettingStepExcelFile.Text = dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[6].Value.ToString();
                //txtRouterIntegrationTestCaseSetting24gSsid.Text = dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[6].Value.ToString();
                //txtRouterIntegrationTestCaseSetting5gSsid.Text = dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[7].Value.ToString();
                //mtbRouterIntegrationTestCaseSettingPcIpAddress.Text = dgvRouterIntegrationTestCaseSettingData.Rows[e.RowIndex].Cells[8].Value.ToString();
            }
        }    
  
        #endregion

        /*============================================================================================*/
        /*=================================== Main Function Area =====================================*/
        /*============================================================================================*/
        #region Main Flow Function

        private void InitRouterIntegrationTestCaseSetting()
        {
            hasDeleteButton = false;
            SetupDataGridViewRouterIntegrationTestCaseSetting();

            /* Load the Data Set last time */
            string filename = System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterIntegrationTestCaseSetting.xml";

            if (File.Exists(filename))
                readXmlRouterIntegrationTestCaseSetting(filename);

            cboxRouterIntegrationTestCaseSettingItemFunctionType.SelectedIndex = 0;
        }

        public void SetupDataGridViewRouterIntegrationTestCaseSetting()
        {
            dgvRouterIntegrationTestCaseSettingData.ColumnCount = 8;
            dgvRouterIntegrationTestCaseSettingData.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvRouterIntegrationTestCaseSettingData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvRouterIntegrationTestCaseSettingData.Name = "Router Integration Test Case Setting";
            //dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgvRouterIntegrationTestCaseSettingData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvRouterIntegrationTestCaseSettingData.Columns[0].Name = "Index";
            dgvRouterIntegrationTestCaseSettingData.Columns[1].Name = "Item Name";
            dgvRouterIntegrationTestCaseSettingData.Columns[2].Name = "Item Function Type";
            dgvRouterIntegrationTestCaseSettingData.Columns[3].Name = "Reset To Defalut After Test";
            dgvRouterIntegrationTestCaseSettingData.Columns[4].Name = "Reboot Dut After Test";
            dgvRouterIntegrationTestCaseSettingData.Columns[5].Name = "Wait Time After Test";
            dgvRouterIntegrationTestCaseSettingData.Columns[6].Name = "Final Excel File";
            dgvRouterIntegrationTestCaseSettingData.Columns[7].Name = "Step Excel File";

            dgvRouterIntegrationTestCaseSettingData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvRouterIntegrationTestCaseSettingData.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvRouterIntegrationTestCaseSettingData.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False; //標題列換行, true -換行, false-不換行
            dgvRouterIntegrationTestCaseSettingData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//標題列置中

            //dgvRouterIntegrationTestCaseSettingData.Columns[0].Width = 120;
            //dgvRouterIntegrationTestCaseSettingData.Columns[1].Width = 120;
            //dgvRouterIntegrationTestCaseSettingData.Columns[2].Width = 120;
            //dgvRouterIntegrationTestCaseSettingData.Columns[3].Width = 80;
            //dgvRouterIntegrationTestCaseSettingData.Columns[4].Width = 80;
            //dgvRouterIntegrationTestCaseSettingData.Columns[5].Width = 80;            
            //dgvRouterIntegrationTestCaseSettingData.Columns[6].Width = 80;
            //dgvRouterIntegrationTestCaseSettingData.Columns[7].Width = 80;            
            //dgvRouterIntegrationTestCaseSettingData.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            //dgvRouterIntegrationTestCaseSettingData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            //dgvRouterIntegrationTestCaseSettingData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; //Use the column width setting
            //dgvRouterIntegrationTestCaseSettingData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            //FillAllColumns(dgvRouterIntegrationTestCaseSettingData, true);
        }

        #endregion

        #region Sub-Function

        private bool CheckParameterRouterIntegrationTestCaseSetting()
        {
            /* Check all data are not empty */
            if (nudRouterIntegrationTestCaseSettingIndex.Value.ToString() == "")
            {
                MessageBox.Show("Index can't be Empty!!!");
                return false;
            }
            
            if (txtRouterIntegrationTestCaseSettingItemName.Text == "")
            {
                MessageBox.Show("Test Case Setting Item Name can't be Empty!!!");
                return false;
            }

            if (cboxRouterIntegrationTestCaseSettingItemFunctionType.Text == "")
            {
                MessageBox.Show("Test Case Setting Item Function Name can't be Empty!!!");
                return false;
            }

            if (nudRouterIntegrationTestCaseSettingWaitTimeAfterTest.Text == "")
            {
                MessageBox.Show("Test Case Setting Wait Time can't be Empty!!!");
                return false;
            }

            //==== Excel File===//
            if (txtRouterIntegrationTestCaseSettingFinalExcelFile.Text == "")
            {
                MessageBox.Show("Test Case Setting Final Excel File can't be Empty!!!");
                return false;
            }

            if (txtRouterIntegrationTestCaseSettingStepExcelFile.Text == "")
            {
                MessageBox.Show("Test Case Setting Final Excel File can't be Empty!!!");
                return false;
            }

            /* Check all files exist */
            if (!File.Exists(txtRouterIntegrationTestCaseSettingFinalExcelFile.Text))
            {
                MessageBox.Show("Test Case Setting Final Excel File doesn't Exist!!!");
                return false;
            }

            if (!File.Exists(txtRouterIntegrationTestCaseSettingStepExcelFile.Text))
            {
                MessageBox.Show("Test Case Setting Step Excel File doesn't Exist!!!");
                return false;
            }

            return true;
        }

        private void RouterIntegrationTestCaseSettingAddSetting()
        {
            /* Check if Delete button exist in the datagridview */
            if (hasDeleteButton)
            {
                hasDeleteButton = false;
                dgvRouterIntegrationTestCaseSettingData.Columns.RemoveAt(0);
                /* Reset the name of btn_TestCondition_Edit as Edit words. */
                btnRouterIntegrationTestCaseSettingEditSetting.Text = "Edit";
            }

            //dgvRouterIntegrationTestCaseSettingData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            //dgvRouterIntegrationTestCaseSettingData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            /* Add data to datagridview */
            // "Index", "Model Name", "Serial Number", "SW Version", "HW Version"
            // "IP Address", "SSID", "Control PC IP" , "Switch Port", "ComPort"
            // "MAC Address", "GUI Script File";
            string[] row = new string[]
            {            
                nudRouterIntegrationTestCaseSettingIndex.Value.ToString(),
                txtRouterIntegrationTestCaseSettingItemName.Text,
                cboxRouterIntegrationTestCaseSettingItemFunctionType.Text,
                chkRouterIntegrationTestCaseSettingDutResetToDefaultAfterTest.Checked? "Y":"N",
                chkRouterIntegrationTestCaseSettingDutRebootAfterTest.Checked? "Y":"N",
                nudRouterIntegrationTestCaseSettingWaitTimeAfterTest.Value.ToString(),
                txtRouterIntegrationTestCaseSettingFinalExcelFile.Text,
                txtRouterIntegrationTestCaseSettingStepExcelFile.Text
            };
            
            dgvRouterIntegrationTestCaseSettingData.Rows.Add(row);

            dgvRouterIntegrationTestCaseSettingData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvRouterIntegrationTestCaseSettingData.AutoResizeColumns();            
        }

        #endregion

        /*============================================================================================*/
        /*=================================== XML Function Area   ====================================*/
        /*============================================================================================*/
        #region        

        public void writeXmlRouterIntegrationTestCaseSetting(string FileName)
        {
            // ToDo: Needs to verify which test item is to be selected..
            int conditionCount = 0;
            string[,] rowdata = new string[dgvRouterIntegrationTestCaseSettingData.RowCount - 1, dgvRouterIntegrationTestCaseSettingData.ColumnCount];

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;

            XmlWriter writer = XmlWriter.Create(FileName, settings);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by the program.");
            writer.WriteStartElement("RouterATETC");
            writer.WriteAttributeString("Item", "Router Integration Test Case Setting");
            
            ///
            /// Write Function Test settings
            /// 
            try
            {
                writer.WriteStartElement("TestCaseSetting");
                for (int i = 0; i < dgvRouterIntegrationTestCaseSettingData.RowCount - 1; i++)
                {
                    conditionCount++;
                    writer.WriteStartElement("Condition_" + (i + 1).ToString());

                    writer.WriteElementString("Index", dgvRouterIntegrationTestCaseSettingData.Rows[i].Cells[0].Value.ToString());
                    writer.WriteElementString("ItemName", dgvRouterIntegrationTestCaseSettingData.Rows[i].Cells[1].Value.ToString());
                    writer.WriteElementString("ItemFunctionName", dgvRouterIntegrationTestCaseSettingData.Rows[i].Cells[2].Value.ToString());
                    writer.WriteElementString("ResetToDefalutAfterTest", dgvRouterIntegrationTestCaseSettingData.Rows[i].Cells[3].Value.ToString());
                    writer.WriteElementString("RebootDutAfterTest", dgvRouterIntegrationTestCaseSettingData.Rows[i].Cells[4].Value.ToString());
                    writer.WriteElementString("WaitTimeAfterTest", dgvRouterIntegrationTestCaseSettingData.Rows[i].Cells[5].Value.ToString());
                    writer.WriteElementString("FinalExcelFile", dgvRouterIntegrationTestCaseSettingData.Rows[i].Cells[6].Value.ToString());
                    writer.WriteElementString("StepExcelFile", dgvRouterIntegrationTestCaseSettingData.Rows[i].Cells[7].Value.ToString());
        
                    writer.WriteEndElement();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Write Test Condition: " + ex.ToString());
            }

            writer.WriteStartElement("Condition_Number");
            writer.WriteElementString("Number", dgvRouterIntegrationTestCaseSettingData.RowCount.ToString());
            writer.WriteEndElement();

            // End of write Function Test
            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public bool readXmlRouterIntegrationTestCaseSetting(string FileName)
        {
            int number = 0;

            //dgvRouterIntegrationTestCaseSettingData.Columns[dgvRouterIntegrationTestCaseSettingData.ColumnCount].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            XmlDocument doc = new XmlDocument();
            doc.Load(FileName);

            XmlNode nodeXml = doc.SelectSingleNode("RouterATETC");
            if (nodeXml == null)
                return false;
            XmlElement element = (XmlElement)nodeXml;
            string strID = element.GetAttribute("Item");
            Debug.WriteLine(strID);
            if (strID.CompareTo("Router Integration Test Case Setting") != 0)
            {
                MessageBox.Show("This XML file is incorrect.", "Error");
                return false;
            }

            ///
            /// Read Function Test configuration settings
            ///

            XmlNode nodeTestConditionModel = doc.SelectSingleNode("/RouterATETC/TestCaseSetting/Condition_Number");
            try
            {
                string Number = nodeTestConditionModel.SelectSingleNode("Number").InnerText;
                Debug.WriteLine("Number: " + Name);
                number = Int32.Parse(Number);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATETC/TestCaseSetting//Condition_Number" + ex);
            }

            for (int i = 0; i < number; i++)
            {
                XmlNode nodeTestCondition = doc.SelectSingleNode("/RouterATETC/TestCaseSetting//Condition_" + (i + 1).ToString());

                try
                {
                    string Index = nodeTestCondition.SelectSingleNode("Index").InnerText;
                    string ItemName = nodeTestCondition.SelectSingleNode("ItemName").InnerText;
                    string ItemFunctionName = nodeTestCondition.SelectSingleNode("ItemFunctionName").InnerText;                    
                    string ResetToDefalutAfterTest = nodeTestCondition.SelectSingleNode("ResetToDefalutAfterTest").InnerText;
                    string RebootDutAfterTest = nodeTestCondition.SelectSingleNode("RebootDutAfterTest").InnerText;
                    string WaitTimeAfterTest = nodeTestCondition.SelectSingleNode("WaitTimeAfterTest").InnerText;
                    string FinalExcelFile = nodeTestCondition.SelectSingleNode("FinalExcelFile").InnerText;
                    string StepExcelFile = nodeTestCondition.SelectSingleNode("StepExcelFile").InnerText;
                   
                    Debug.WriteLine("Index: " + Index);
                    Debug.WriteLine("ItemName: " + ItemName);
                    Debug.WriteLine("ItemFunctionName: " + ItemFunctionName);
                    Debug.WriteLine("ResetToDefalutAfterTest: " + ResetToDefalutAfterTest);
                    Debug.WriteLine("RebootDutAfterTest: " + RebootDutAfterTest);
                    Debug.WriteLine("WaitTimeAfterTest: " + WaitTimeAfterTest);
                    Debug.WriteLine("FinalExcelFile: " + FinalExcelFile);
                    Debug.WriteLine("StepExcelFile: " + StepExcelFile);

                    string[] data = new string[] { Index, ItemName, ItemFunctionName, ResetToDefalutAfterTest, RebootDutAfterTest, WaitTimeAfterTest, FinalExcelFile, StepExcelFile};
                    dgvRouterIntegrationTestCaseSettingData.Rows.Add(data);

                }
                catch (Exception ex)
                {
                    Debug.WriteLine("/RouterATETC/TestCaseSetting//Condition_Read " + ex);
                }
            }

            // End of Read Test Condition configuration settings
            //FillAllColumns(dgvRouterIntegrationTestCaseSettingData, true);
            return true;
        }

        #endregion
        
        /*============================================================================================*/
        /*======================================= The End   ==========================================*/
        /*============================================================================================*/


    }
}
