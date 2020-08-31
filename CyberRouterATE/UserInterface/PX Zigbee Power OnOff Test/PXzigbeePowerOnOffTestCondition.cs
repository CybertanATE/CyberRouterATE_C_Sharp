///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Sally Lee.
///  File           : PXzigbeePowerOnOffTestCondition.cs
///  Update         : 2020-08-11
///  Modified       : 2020-08-11 Initial version  
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
        bool bHasDeleteButton_pxZigbeePowerOnOff = false;

        //**********************************************************************************//
        //------------------- PX Zigbee Power OnOff Test Condition Event -------------------//
        //**********************************************************************************//
#region PX Zigbee Power OnOff Test Condition Event
        private void cboxPXzigbeePowerOnOffTestConditionAction1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string pxZigbeePowerOnHeader  = "<PX Zigbee Power On Command>";
            string pxZigbeePowerOffHeader = "<PX Zigbee Power Off Command>";
            string getLightStatusHeader   = "<Get Light Status Command>";

            labelPXzigbeePowerOnOffTestConditionCMD.Text = "";
        }

        private void btnPXzigbeePowerOnOffTestConditionAddSetting_Click(object sender, EventArgs e)
        {
            /* Check if Delete button exist in the datagridview */
            if (bHasDeleteButton_pxZigbeePowerOnOff)
            {
                bHasDeleteButton_pxZigbeePowerOnOff = false;
                dgvPXzigbeePowerOnOffTestConditionData.Columns.RemoveAt(0);
                /* Reset the name of btn_TestCondition_Edit as Edit words. */
                btnPXzigbeePowerOnOffTestConditionEditSetting.Text = "Edit";
            }



            string ModelName   = txtPXzigbeePowerOnOffTestConditionModelName.Text;
            string DeviceIP    = txtPXzigbeePowerOnOffTestConditionIP.Text;
            string NodeID      = txtPXzigbeePowerOnOffTestConditionNodeID.Text;
            string Action1     = cboxPXzigbeePowerOnOffTestConditionAction1.SelectedItem.ToString();
            string SleepTimer1 = nudPXzigbeePowerOnOffTestConditionAction1SleepTimer.Value.ToString();
            string Action2     = cboxPXzigbeePowerOnOffTestConditionAction2.SelectedItem.ToString();
            string SleepTimer2 = nudPXzigbeePowerOnOffTestConditionAction2SleepTimer.Value.ToString();
            string Action3     = cboxPXzigbeePowerOnOffTestConditionAction3.SelectedItem.ToString();
            string SleepTimer3 = nudPXzigbeePowerOnOffTestConditionAction3SleepTimer.Value.ToString();
            string Action4     = cboxPXzigbeePowerOnOffTestConditionAction4.SelectedItem.ToString();
            string SleepTimer4 = nudPXzigbeePowerOnOffTestConditionAction4SleepTimer.Value.ToString();
            string MqttCmdPath = txtPXzigbeePowerOnOffTestConditionMqttCmdPath.Text;

            /* Add data to datagridview */
            string[] data = new string[] 
            {
                (ModelName == "")? "X":ModelName,
                DeviceIP,
                NodeID,
                Action1,
                SleepTimer1,
                Action2,
                SleepTimer2,
                Action3,
                SleepTimer3,
                Action4,
                SleepTimer4,
                (MqttCmdPath == "")? @"C:\Program Files\Project X\mos158":MqttCmdPath
            };

            dgvPXzigbeePowerOnOffTestConditionData.Rows.Add(data);
            dgvPXzigbeePowerOnOffTestConditionData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvPXzigbeePowerOnOffTestConditionData.AutoResizeColumns();
            dgvPXzigbeePowerOnOffTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
        }

        private void btnPXzigbeePowerOnOffTestConditionEditSetting_Click(object sender, EventArgs e)
        {
            /* James: To prevent one cell is to be inserted more time when Edit button clicked once again. */
            if (bHasDeleteButton_pxZigbeePowerOnOff == false)
            {
                bHasDeleteButton_pxZigbeePowerOnOff = true;
                btnPXzigbeePowerOnOffTestConditionEditSetting.Text = "Cancel";
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.Width = 66;
                btn.HeaderText = "Action";
                btn.Text = "Delete";
                btn.Name = "Action";
                btn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                btn.UseColumnTextForButtonValue = true;

                dgvPXzigbeePowerOnOffTestConditionData.Columns.Insert(0, btn);
            }
            else
            {
                btnPXzigbeePowerOnOffTestConditionEditSetting.Text = "Edit";
                bHasDeleteButton_pxZigbeePowerOnOff = false;
                dgvPXzigbeePowerOnOffTestConditionData.Columns.Remove("Action");
            }
        }

        private void btnPXzigbeePowerOnOffTestConditionSaveSetting_Click(object sender, EventArgs e)
        {

        }

        private void btnPXzigbeePowerOnOffTestConditionLoadSetting_Click(object sender, EventArgs e)
        {

        }

        private void dgvPXzigbeePowerOnOffTestConditionData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            /* Adding a condition to prevent exception error when remove null cell. */
            if (dgvPXzigbeePowerOnOffTestConditionData.Columns[e.ColumnIndex].Name == "Action" && bHasDeleteButton_pxZigbeePowerOnOff == true && dgvPXzigbeePowerOnOffTestConditionData.Rows[e.RowIndex].Cells[0].Value != null)
            {
                dgvPXzigbeePowerOnOffTestConditionData.Rows.RemoveAt(e.RowIndex);
                if (e.ColumnIndex == 0 && (String)dgvPXzigbeePowerOnOffTestConditionData.Rows[0].Cells[0].Value == String.Empty)
                    dgvPXzigbeePowerOnOffTestConditionData.Columns.RemoveAt(0);
            }
        }
#endregion //-- PX Zigbee Power OnOff Test Condition Event




        //**********************************************************************************//
        //------------------- PX Zigbee Power OnOff Test Condition Module ------------------//
        //**********************************************************************************//
#region PX Zigbee Power OnOff Test Condition Module
        private void InitPXzigbeePowerOnOffTestCondition()
        {
            /* indicate Delete button exist in the datagridview column 0*/
            bHasDeleteButton_pxZigbeePowerOnOff = false;
            InitPXzigbeePowerOnOffDataGridView();
        }

        public void InitPXzigbeePowerOnOffDataGridView()
        {
            dgvPXzigbeePowerOnOffTestConditionData.ColumnCount = 12;
            dgvPXzigbeePowerOnOffTestConditionData.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvPXzigbeePowerOnOffTestConditionData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvPXzigbeePowerOnOffTestConditionData.Name = "PX Zigbee Power On/Off Test Condition Setting";
            dgvPXzigbeePowerOnOffTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvPXzigbeePowerOnOffTestConditionData.Columns[0].Name = "Mode Name";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[1].Name = "IP";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[2].Name = "Node ID";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[3].Name = "Action 1";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[4].Name = "Sleep Timer 1";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[5].Name = "Action 2";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[6].Name = "Sleep Timer 2";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[7].Name = "Action 3";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[8].Name = "Sleep Timer 3";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[9].Name = "Action 4";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[10].Name = "Sleep Timer 4";
            dgvPXzigbeePowerOnOffTestConditionData.Columns[11].Name = "MQTT Path";

            dgvPXzigbeePowerOnOffTestConditionData.DefaultCellStyle.Alignment              = DataGridViewContentAlignment.MiddleCenter;
            dgvPXzigbeePowerOnOffTestConditionData.DefaultCellStyle.WrapMode               = DataGridViewTriState.True;
            dgvPXzigbeePowerOnOffTestConditionData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; //標題列置中
            dgvPXzigbeePowerOnOffTestConditionData.ColumnHeadersDefaultCellStyle.WrapMode  = DataGridViewTriState.True;  //標題列換行, True -換行, False-不換行

            dgvPXzigbeePowerOnOffTestConditionData.Columns[0].Width = 100;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[1].Width = 60;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[2].Width = 100;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[3].Width = 150;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[4].Width = 80;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[5].Width = 150;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[6].Width = 80;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[7].Width = 150;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[8].Width = 80;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[9].Width = 150;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[10].Width = 80;
            dgvPXzigbeePowerOnOffTestConditionData.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;


            //dgvPXzigbeePowerOnOffTestConditionData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dgvPXzigbeePowerOnOffTestConditionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
        }
#endregion //-- PX Zigbee Power OnOff Test Condition Module
    }
}
