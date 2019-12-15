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
               

        private void InitPowerOnOffTestCondition()
        {
            /* indicate Delete button exist in the datagridview column 0*/
            hasDeleteButton = false;
            SetupPowerOnOffDataGridView();            
        }

        public void SetupPowerOnOffDataGridView()
        {
            dgvPowerOnOffTestConditionData.ColumnCount = 7;
            dgvPowerOnOffTestConditionData.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvPowerOnOffTestConditionData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvPowerOnOffTestConditionData.Name = "Power OnOff Test Condition Setting";
            //dgvPowerOnOffTestConfitionData.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgvPowerOnOffTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvPowerOnOffTestConditionData.Columns[0].Name = "Power Port";
            dgvPowerOnOffTestConditionData.Columns[1].Name = "Model Name";
            dgvPowerOnOffTestConditionData.Columns[2].Name = "Action";
            dgvPowerOnOffTestConditionData.Columns[3].Name = "Parameter 1";
            dgvPowerOnOffTestConditionData.Columns[4].Name = "Parameter 2";
            dgvPowerOnOffTestConditionData.Columns[5].Name = "Power On Time";
            dgvPowerOnOffTestConditionData.Columns[6].Name = "Power Off Time";

            dgvPowerOnOffTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPowerOnOffTestConditionData.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvPowerOnOffTestConditionData.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False; //標題列換行, true -換行, false-不換行
            dgvPowerOnOffTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//標題列置中

            dgvPowerOnOffTestConditionData.Columns[0].Width = 80;
            dgvPowerOnOffTestConditionData.Columns[1].Width = 120;
            dgvPowerOnOffTestConditionData.Columns[2].Width = 100;
            dgvPowerOnOffTestConditionData.Columns[3].Width = 120;
            dgvPowerOnOffTestConditionData.Columns[4].Width = 120;
            dgvPowerOnOffTestConditionData.Columns[5].Width = 100;
            dgvPowerOnOffTestConditionData.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dgvPowerOnOffTestConditionData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            //dgvPowerOnOffTestConfitionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; //Use the column width setting
            dgvPowerOnOffTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            /*
             http://blog.csdn.net/alisa525/article/details/7556771
             // 設定包括Header和所有儲存格的列寬自動調整
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;  //設置列標題不換行
             // 設定不包括Header所有儲存格的行高自動調整
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;  //AllCells;設定包括Header和所有儲存格的行高自動調整            
             */
        }

        private void btnPowerOnOffTestConditionAddSetting_Click(object sender, EventArgs e)
        {  
            /* Check if Delete button exist in the datagridview */
            if (hasDeleteButton)
            {
                hasDeleteButton = false;
                dgvPowerOnOffTestConditionData.Columns.RemoveAt(0);
                /* Reset the name of btn_TestCondition_Edit as Edit words. */
                btnRvRTestConditionEditSetting.Text = "Edit";
            }
                        
            /* Add data to datagridview */
            string[] data = new string[] {
                nudPowerOnOffTestConditionPowerPort.Value.ToString(), //Power port
                (txtPowerOnOffTestConditionModelName.Text == "")? "X":txtPowerOnOffTestConditionModelName.Text, //Model Name
                cboxPowerOnOffTestConditionAction.SelectedItem.ToString(), //Action
                (txtPowerOnOffTestConditionParameter1.Text =="")? "X":txtPowerOnOffTestConditionParameter1.Text, //parameter 1
                (txtPowerOnOffTestConditionParameter2.Text == "")? "X":txtPowerOnOffTestConditionParameter2.Text, //Parameter 2
                nudPowerOnOffTestConditionPowerOnTime.Value.ToString(), //Power on time
                nudPowerOnOffTestConditionPowerOffTime.Value.ToString()}; //Power off time

            //string[] data = new string[] { "1", "2", "3", "4", "5", "6", "7" };
                //"1", //Power port
                //txtPowerOnOffTestConditionModelName.Text, //Model Name
                //cboxPowerOnOffTestConditionAction.SelectedItem.ToString(), //Action
                //txtPowerOnOffTestConditionParameter1.Text, //parameter 1
                //txtPowerOnOffTestConditionParameter2.Text, //Parameter 2
                //nudPowerOnOffTestConditionPowerOnTime.Value.ToString(), //Power on time
                //nudPowerOnOffTestConditionPowerOffTime.Value.ToString()}; //Power off time

            
            dgvPowerOnOffTestConditionData.Rows.Add(data) ;
            dgvPowerOnOffTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
           
            dgvPowerOnOffTestConditionData.AutoResizeColumns();
            dgvPowerOnOffTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;        
        }

        private void btnPowerOnOffTestConditionEditSetting_Click(object sender, EventArgs e)
        {
            /* James: To prevent one cell is to be inserted more time when Edit button clicked once again. */
            if (hasDeleteButton == false)
            {
                hasDeleteButton = true;
                btnPowerOnOffTestConditionEditSetting.Text = "Cancel";
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.Width = 66;
                btn.HeaderText = "Action";
                btn.Text = "Delete";
                btn.Name = "Action";
                btn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                btn.UseColumnTextForButtonValue = true;

                dgvPowerOnOffTestConditionData.Columns.Insert(0, btn);
            }
            else
            {
                btnPowerOnOffTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvPowerOnOffTestConditionData.Columns.Remove("Action");
            }
        }

        private void btnPowerOnOffTestConditionSaveSetting_Click(object sender, EventArgs e)
        {
            if (dgvPowerOnOffTestConditionData.RowCount <= 1)
            {
                MessageBox.Show("Add Condition First!!", "Warning");
                return;
            }

            string filename = "PowerOnOffTestCondition";

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
                writeXmlPowerOnOffTestCondition(saveFileDialog1.FileName);
            }
        }

        private void btnPowerOnOffTestConditionLoadSetting_Click(object sender, EventArgs e)
        {
            if (btnPowerOnOffTestConditionEditSetting.Text.ToLower() == "cancel")
            {
                btnPowerOnOffTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvPowerOnOffTestConditionData.Columns.Remove("Action");
            }

            if (dgvPowerOnOffTestConditionData.RowCount > 1)
            {
                if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                    return;
                else
                {
                    DataTable dt = (DataTable)dgvPowerOnOffTestConditionData.DataSource;
                    dgvPowerOnOffTestConditionData.Rows.Clear();
                    dgvPowerOnOffTestConditionData.DataSource = dt;
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
                readXmlPowerOnOffTestCondition(openFileDialog1.FileName);
            }
        }

        private void dgvPowerOnOffTestConfition_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            /* James modified: Adding a condition to prevent exception error when remove null cell. */
            if (dgvPowerOnOffTestConditionData.Columns[e.ColumnIndex].Name == "Action" && hasDeleteButton == true && dgvPowerOnOffTestConditionData.Rows[e.RowIndex].Cells[0].Value != null)
            {
                dgvPowerOnOffTestConditionData.Rows.RemoveAt(e.RowIndex);
                if (e.ColumnIndex == 0 && (String)dgvPowerOnOffTestConditionData.Rows[0].Cells[0].Value == String.Empty)
                    dgvPowerOnOffTestConditionData.Columns.RemoveAt(0);
            }
        }

        public void writeXmlPowerOnOffTestCondition(string FileName)
        {
            // ToDo: Needs to verify which test item is to be selected..

            string[,] rowdata = new string[dgvPowerOnOffTestConditionData.RowCount - 1, 7];

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;

            XmlWriter writer = XmlWriter.Create(FileName, settings);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by the program.");
            writer.WriteStartElement("CyberRouterATETC");
            writer.WriteAttributeString("Item", "Power OnOff Test Condition");

            ///
            /// Write Function Test settings
            /// 
            writer.WriteStartElement("TestCondition");
            for (int i = 0; i < dgvPowerOnOffTestConditionData.RowCount - 1; i++)
            {
                writer.WriteStartElement("Condition_" + (i + 1).ToString());

                writer.WriteElementString("Power_Port", rowdata[i, 0] = dgvPowerOnOffTestConditionData.Rows[i].Cells[0].Value.ToString());
                writer.WriteElementString("Model_Name", rowdata[i, 1] = dgvPowerOnOffTestConditionData.Rows[i].Cells[1].Value.ToString());
                writer.WriteElementString("Action", rowdata[i, 2] = dgvPowerOnOffTestConditionData.Rows[i].Cells[2].Value.ToString());
                writer.WriteElementString("Parameter_1", rowdata[i, 3] = dgvPowerOnOffTestConditionData.Rows[i].Cells[3].Value.ToString());
                writer.WriteElementString("Parameter_2", rowdata[i, 4] = dgvPowerOnOffTestConditionData.Rows[i].Cells[4].Value.ToString());
                writer.WriteElementString("Power_On_Time", rowdata[i, 5] = dgvPowerOnOffTestConditionData.Rows[i].Cells[5].Value.ToString());
                writer.WriteElementString("Power_Off_Time", rowdata[i, 6] = dgvPowerOnOffTestConditionData.Rows[i].Cells[6].Value.ToString());

                writer.WriteEndElement();
            }

            writer.WriteStartElement("Condition_Number");
            writer.WriteElementString("Number", dgvPowerOnOffTestConditionData.RowCount.ToString());
            writer.WriteEndElement();

            // End of write Function Test
            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public bool readXmlPowerOnOffTestCondition(string FileName)
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
            if (strID.CompareTo("Power OnOff Test Condition") != 0)
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
                    string PowerPort = nodeTestCondition.SelectSingleNode("Power_Port").InnerText;
                    string ModelName = nodeTestCondition.SelectSingleNode("Model_Name").InnerText;
                    string Action = nodeTestCondition.SelectSingleNode("Action").InnerText;
                    string Parameter1 = nodeTestCondition.SelectSingleNode("Parameter_1").InnerText;
                    string Parameter2 = nodeTestCondition.SelectSingleNode("Parameter_2").InnerText;
                    string PowerOnTime = nodeTestCondition.SelectSingleNode("Power_On_Time").InnerText;
                    string PowerOffTime = nodeTestCondition.SelectSingleNode("Power_Off_Time").InnerText;

                    Debug.WriteLine("Power_Port: " + PowerPort);
                    Debug.WriteLine("Model_Name: " + ModelName);
                    Debug.WriteLine("Action: " + Action);
                    Debug.WriteLine("Parameter_1: " + Parameter1);
                    Debug.WriteLine("Parameter_2: " + Parameter2);
                    Debug.WriteLine("Power_On_Time: " + PowerOnTime);
                    Debug.WriteLine("Power_Off_Time: " + PowerOffTime);                    

                    string[] data = new string[] { PowerPort, ModelName, Action, Parameter1, Parameter2, PowerOnTime, PowerOffTime};
                    dgvPowerOnOffTestConditionData.Rows.Add(data);
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
