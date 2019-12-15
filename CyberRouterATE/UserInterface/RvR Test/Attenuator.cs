using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RvRTest
{
    public partial class Attenuator : Form
    {
        public decimal atteuator_no_2_4G = 0;
        public decimal atteuator_no_5G = 0;
        public string[] atteuator_ip_2_4G;
        public string[] attenuator_ip_5G;
        public decimal[] attenuator_value_2_4G;
        public decimal[] attenuator_value_5G;
        public int gpib_interface_2_4G;
        public int gpib_interface_5G;


        public Attenuator()
        {
            InitializeComponent();
            //ToggleAtteuatorValue(false);
        }

        private void ToggleAtteuatorValue(bool Toggle)
        {
            /* Prevent Attenuator valure of 2.4 G and 5G to be changed. */
            nudAtteuation_2_4G_X1.Enabled = Toggle;
            nudAtteuation_2_4G_X2.Enabled = Toggle;
            nudAtteuation_2_4G_X3.Enabled = Toggle;
            nudAtteuation_2_4G_X4.Enabled = Toggle;
            nudAtteuation_2_4G_X5.Enabled = Toggle;
            nudAtteuation_2_4G_X6.Enabled = Toggle;
            nudAtteuation_2_4G_X7.Enabled = Toggle;
            nudAtteuation_2_4G_X8.Enabled = Toggle;

            nudAtteuation_5G_X1.Enabled = Toggle;
            nudAtteuation_5G_X2.Enabled = Toggle;
            nudAtteuation_5G_X3.Enabled = Toggle;
            nudAtteuation_5G_X4.Enabled = Toggle;
            nudAtteuation_5G_X5.Enabled = Toggle;
            nudAtteuation_5G_X6.Enabled = Toggle;
            nudAtteuation_5G_X7.Enabled = Toggle;
            nudAtteuation_5G_X8.Enabled = Toggle;
        }

        private void lab_Attenuation_Add_2_4_Click(object sender, EventArgs e)
        {
            /* Check if IP address exist in the listbox */
            foreach (string item in lbox_AtteuationSetting_GPIBIP_2_4G.Items)
            {
                if (nud_Atteuation_GPIBIPAdderss_2_4G.Value.ToString() == item) 
                    return;
            }

            lbox_AtteuationSetting_GPIBIP_2_4G.Items.Add(nud_Atteuation_GPIBIPAdderss_2_4G.Value.ToString());
        }

        private void lab_Attenuation_Add_5G_Click(object sender, EventArgs e)
        {
            /* Check if IP address exist in the listbox */
            foreach (string item in lbox_AtteuationSetting_GPIBIP_5G.Items) 
            {                
                if (nud_Attenuation_GPIBIPAddress_5G.Value.ToString() == item)
                    return;
            }

            lbox_AtteuationSetting_GPIBIP_5G.Items.Add(nud_Attenuation_GPIBIPAddress_5G.Value.ToString());
        }

        private void btn_AttenuatorSetting_OK_Click(object sender, EventArgs e)
        {
            if (nud_AtteuatorNumber_2_4G.Value == 0 && nud_AtteuatorNumber_5G.Value == 0)
            {
                MessageBox.Show("Attenuator Number is Zero!!! No attenuation was selected.","Warning");
                return;
            }
            if (nud_AtteuatorNumber_2_4G.Value != 0)
            {
                if (nud_AtteuatorNumber_2_4G.Value != lbox_AtteuationSetting_GPIBIP_2_4G.Items.Count)
                {
                    MessageBox.Show("2.4G : Attenuator number is not equal to the ip address number!!! ", "Warning");
                    return;
                }

                gpib_interface_2_4G = (int) nud_GPIB_Interface_2_4G.Value;
                atteuator_no_2_4G = nud_AtteuatorNumber_2_4G.Value;
                atteuator_ip_2_4G = new string[(int)atteuator_no_2_4G];
                int i = 0 ;
                foreach (string item in lbox_AtteuationSetting_GPIBIP_2_4G.Items)
                {
                    atteuator_ip_2_4G[i++] = item;                    
                }
                
                

                attenuator_value_2_4G = new decimal[] 
                    {nudAtteuation_2_4G_X1.Value, nudAtteuation_2_4G_X2.Value,
                    nudAtteuation_2_4G_X3.Value, nudAtteuation_2_4G_X4.Value,
                    nudAtteuation_2_4G_X5.Value, nudAtteuation_2_4G_X6.Value,
                    nudAtteuation_2_4G_X7.Value, nudAtteuation_2_4G_X7.Value};
            }
            
            if (nud_AtteuatorNumber_5G.Value != 0)
            {
                if (nud_AtteuatorNumber_5G.Value != lbox_AtteuationSetting_GPIBIP_5G.Items.Count)
                {
                    MessageBox.Show("5G : Attenuator number is not equal to the ip address number!!! ", "Warning");
                    return;
                }

                gpib_interface_5G = (int)nud_GPIB_Interface_5G.Value;
                atteuator_no_5G = nud_AtteuatorNumber_5G.Value;
                attenuator_ip_5G = new string[(int)atteuator_no_5G];
                int i = 0;
                foreach (string item in lbox_AtteuationSetting_GPIBIP_5G.Items)
                {
                    attenuator_ip_5G[i++] = item;
                }
                
                attenuator_value_5G = new decimal[] 
                    {nudAtteuation_5G_X1.Value, nudAtteuation_5G_X2.Value,
                    nudAtteuation_5G_X3.Value, nudAtteuation_5G_X4.Value,
                    nudAtteuation_5G_X5.Value, nudAtteuation_5G_X6.Value,
                    nudAtteuation_5G_X7.Value, nudAtteuation_5G_X8.Value};
            }

            this.Close();

        }

        private void nudAtteuation_2_4G_X1_ValueChanged(object sender, EventArgs e)
        {
            nudAtteuation_2_4G_X2.Minimum = nudAtteuation_2_4G_X1.Value;
            nudAtteuation_2_4G_X3.Minimum = nudAtteuation_2_4G_X1.Value;
            nudAtteuation_2_4G_X4.Minimum = nudAtteuation_2_4G_X1.Value;
        }

        private void nudAtteuation_2_4G_X2_ValueChanged(object sender, EventArgs e)
        {
            nudAtteuation_2_4G_X3.Minimum = nudAtteuation_2_4G_X2.Value;
            nudAtteuation_2_4G_X4.Minimum = nudAtteuation_2_4G_X2.Value;
        }

        private void nudAtteuation_2_4G_X3_ValueChanged(object sender, EventArgs e)
        {
            nudAtteuation_2_4G_X4.Minimum = nudAtteuation_2_4G_X3.Value;
        }

        private void nudAtteuation_2_4G_X5_ValueChanged(object sender, EventArgs e)
        {
            nudAtteuation_2_4G_X6.Minimum = nudAtteuation_2_4G_X5.Value;
            nudAtteuation_2_4G_X7.Minimum = nudAtteuation_2_4G_X5.Value;
            nudAtteuation_2_4G_X8.Minimum = nudAtteuation_2_4G_X5.Value;
        }

        private void nudAtteuation_2_4G_X6_ValueChanged(object sender, EventArgs e)
        {
            nudAtteuation_2_4G_X7.Minimum = nudAtteuation_2_4G_X6.Value;
            nudAtteuation_2_4G_X8.Minimum = nudAtteuation_2_4G_X6.Value;
        }

        private void nudAtteuation_2_4G_X7_ValueChanged(object sender, EventArgs e)
        {
            nudAtteuation_2_4G_X8.Minimum = nudAtteuation_2_4G_X7.Value;
        }

        private void nudAtteuation_5G_X1_ValueChanged(object sender, EventArgs e)
        {
            nudAtteuation_5G_X2.Minimum = nudAtteuation_5G_X1.Value;
            nudAtteuation_5G_X3.Minimum = nudAtteuation_5G_X1.Value;
            nudAtteuation_5G_X4.Minimum = nudAtteuation_5G_X1.Value;
        }

        private void nudAtteuation_5G_X2_ValueChanged(object sender, EventArgs e)
        {
            nudAtteuation_5G_X3.Minimum = nudAtteuation_5G_X2.Value;
            nudAtteuation_5G_X4.Minimum = nudAtteuation_5G_X2.Value;
        }

        private void nudAtteuation_5G_X3_ValueChanged(object sender, EventArgs e)
        {            
            nudAtteuation_5G_X4.Minimum = nudAtteuation_5G_X3.Value;
        }

        private void nudAtteuation_5G_X5_ValueChanged(object sender, EventArgs e)
        {
            nudAtteuation_5G_X6.Minimum = nudAtteuation_5G_X5.Value;
            nudAtteuation_5G_X7.Minimum = nudAtteuation_5G_X5.Value;
            nudAtteuation_5G_X8.Minimum = nudAtteuation_5G_X5.Value;            
        }

        private void nudAtteuation_5G_X6_ValueChanged(object sender, EventArgs e)
        {
            nudAtteuation_5G_X7.Minimum = nudAtteuation_5G_X6.Value;
            nudAtteuation_5G_X8.Minimum = nudAtteuation_5G_X6.Value;
        }

        private void nudAtteuation_5G_X7_ValueChanged(object sender, EventArgs e)
        {            
            nudAtteuation_5G_X8.Minimum = nudAtteuation_5G_X7.Value;
        }
    }
}
