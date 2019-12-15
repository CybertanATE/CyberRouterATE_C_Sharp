///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RvRFunctionTest.cs
///  Update         : 2016-07-25
///  Description    : Main function
///  Modified       : 2016-07-25 Initial version  
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
using Ivi.Visa.Interop;
//using NationalInstruments.VisaNS;
using System.Diagnostics;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        private void btn_AttenuationSetting_GPIB_Info_Click(object sender, EventArgs e)
        {
            string[] res = FindGPIBResource();

            foreach (string s in res)
            {
                txt_AttenuationSetting_Information.AppendText(s + Environment.NewLine);
            }
        }             

        private string[] FindGPIBResource()
        {
            ResourceManager rm = new ResourceManager();
            string filter = "?*";
            string[] resource = rm.FindRsrc(filter);           
             
            return resource;
        }

        private void labRvRAttenuationAdd24G_Click(object sender, EventArgs e)
        {
            /* Check if IP address exist in the listbox */
            // Ignore below condition by James for testing
            /*
            foreach (string item in lbox_AtteuationSetting_GPIBIP_2_4G.Items)
            {
                if (nud_Atteuation_GPIBIPAdderss_2_4G.Value.ToString() == item) 
                    return;
            }
            */
            lboxRvRAtteuationSettingGPIBIP24G.Items.Add(nudRvRAtteuationGPIBIPAdderss24G.Value.ToString());
        }
        
        private void labRvRAttenuationDel24G_Click(object sender, EventArgs e)
        {
            int index = lboxRvRAtteuationSettingGPIBIP24G.SelectedIndex;
            if (index >= 0)
            {
                lboxRvRAtteuationSettingGPIBIP24G.Items.RemoveAt(index);
            }     
        }

        private void labRvRAttenuationAdd5G_Click(object sender, EventArgs e)
        {
            /* Check if IP address exist in the listbox */
            // Ignore below condition by James for testing
            /*
            foreach (string item in lbox_AtteuationSetting_GPIBIP_5G.Items)
            {
                if (nud_Attenuation_GPIBIPAddress_5G.Value.ToString() == item)
                    return;
            }
            */
            lboxRvRAtteuationSettingGPIBIP5G.Items.Add(nudRvRAttenuationGPIBIPAddress5G.Value.ToString());
        }
        
        private void labRvRAttenuationDel5G_Click(object sender, EventArgs e)
        {
            int index = lboxRvRAtteuationSettingGPIBIP5G.SelectedIndex;
            if (index >= 0)
            {
                lboxRvRAtteuationSettingGPIBIP5G.Items.RemoveAt(index);
            } 
        }

        private void nudRvRAtteuation24GX1_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation24GX2.Minimum = nudRvRAtteuation24GX1.Value;
            nudRvRAtteuation24GX3.Minimum = nudRvRAtteuation24GX1.Value;
            nudRvRAtteuation24GX4.Minimum = nudRvRAtteuation24GX1.Value;
        }

        private void nudRvRAtteuation24GX2_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation24GX3.Minimum = nudRvRAtteuation24GX2.Value;
            nudRvRAtteuation24GX4.Minimum = nudRvRAtteuation24GX2.Value;
        }

        private void nudRvRAtteuation24GX3_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation24GX4.Minimum = nudRvRAtteuation24GX3.Value;
        }

        private void nudRvRAtteuation24GX5_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation24GX6.Minimum = nudRvRAtteuation24GX5.Value;
            nudRvRAtteuation24GX7.Minimum = nudRvRAtteuation24GX5.Value;
            nudRvRAtteuation24GX8.Minimum = nudRvRAtteuation24GX5.Value;
        }

        private void nudRvRAtteuation24GX6_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation24GX7.Minimum = nudRvRAtteuation24GX6.Value;
            nudRvRAtteuation24GX8.Minimum = nudRvRAtteuation24GX6.Value;
        }

        private void nudRvRAtteuation24GX7_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation24GX8.Minimum = nudRvRAtteuation24GX7.Value;
        }

        private void nudRvRAtteuation5GX1_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation5GX2.Minimum = nudRvRAtteuation5GX1.Value;
            nudRvRAtteuation5GX3.Minimum = nudRvRAtteuation5GX1.Value;
            nudRvRAtteuation5GX4.Minimum = nudRvRAtteuation5GX1.Value;
        }

        private void nudRvRAtteuation5GX2_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation5GX3.Minimum = nudRvRAtteuation5GX2.Value;
            nudRvRAtteuation5GX4.Minimum = nudRvRAtteuation5GX2.Value;
        }

        private void nudRvRAtteuation5GX3_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation5GX4.Minimum = nudRvRAtteuation5GX3.Value;
        }

        private void nudRvRAtteuation5GX5_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation5GX6.Minimum = nudRvRAtteuation5GX5.Value;
            nudRvRAtteuation5GX7.Minimum = nudRvRAtteuation5GX5.Value;
            nudRvRAtteuation5GX8.Minimum = nudRvRAtteuation5GX5.Value;
        }

        private void nudRvRAtteuation5GX6_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation5GX7.Minimum = nudRvRAtteuation5GX6.Value;
            nudRvRAtteuation5GX8.Minimum = nudRvRAtteuation5GX6.Value;
        }

        private void nudRvRAtteuation5GX7_ValueChanged(object sender, EventArgs e)
        {
            nudRvRAtteuation5GX8.Minimum = nudRvRAtteuation5GX7.Value;
        }
              
        private void ToggleAttenuatorSettingControl(bool Toggle)
        {
             /* Test Condition */
             rbtnRvRTestCondition24G.Enabled = Toggle;
             rbtnRvRTestCondition5G.Enabled = Toggle;
             cboxRvRTestConditionWirelessMode.Enabled = Toggle;
             cbox_Model_Name.Enabled = Toggle;
             cboxRvRTestConditionWirelessChannel.Enabled = Toggle;
             txt_TestCondition_SSID.Enabled = Toggle;
             nud_TestCondition_AtteuationMin.Enabled = Toggle;
             nud_TestCondition_AtteuationMax.Enabled = Toggle;
             nud_TestCondition_AtteuationStep.Enabled = Toggle;
             txtRvRTestConditionTxTst.Enabled = Toggle;
             txtRvRTestConditionRxTst.Enabled = Toggle;
             txtRvRTestConditionBiTst.Enabled = Toggle;
             btnRvRTestConditionAddSetting.Enabled = Toggle;             
             btnRvRTestConditionEditSetting.Enabled = Toggle;

             /* Function Test */
             cbox_Model_Name.Enabled = Toggle;
             txt_Model_SerialNumber.Enabled = Toggle;
             txt_Model_HWVersion.Enabled = Toggle;
             txt_Model_SWVersion.Enabled = Toggle;
             txt_RouterSetting_UserName.Enabled = Toggle;
             txt_RouterSetting_Password.Enabled = Toggle;
             txt_RouterSetting_DUTIPAddress.Enabled = Toggle;
             txt_ClientSetting_24gIPaddress.Enabled = Toggle;
             txt_ClientSetting_5gIPaddress.Enabled = Toggle;
             txt_ChariotSetting_ChariotFolder.Enabled = Toggle;

             btnRvRFunctionTestRun.Text = Toggle ? "Run" : "Stop";
             Debug.WriteLine(btnRvRFunctionTestRun.Text);
        }
        
    }//End of class
}
