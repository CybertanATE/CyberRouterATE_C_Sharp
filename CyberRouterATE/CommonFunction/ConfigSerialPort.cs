///---------------------------------------------------------------------------------------
///  This code was created by CyberTan James Chu.
///  File           : ConfigSerialPort.cs
///  Update         : 2014-07-22
///  Version        : 1.0.140722
///  Description    : 
///  Modified       : 2014-07-22 Initial version
///---------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
//using UART;
using ComportClass;

namespace CyberRouterATE
{
    public partial class ConfigSerialPort : Form
    {
        Comport comport = null;

        public ConfigSerialPort()
        {
            InitializeComponent();
        }

        private void ConfigSerialPort_Load(object sender, EventArgs e)
        {
            /* Get a list of serial port names */
            string[] Ports = SerialPort.GetPortNames();
            cbPort.SelectedIndex = -1;
            foreach (string port in Ports)
            {
                cbPort.Items.Add(port);
                // set index value to 0 if serial device finding
                cbPort.SelectedIndex = 0;
            }

            /* Set the default data bit */
            cbData.SelectedIndex = 1;

            /* Set the default parity */
            cbParity.SelectedIndex = 0;

            /* Set the default stop bit */
            cbStop.SelectedIndex = 0;

            /* Set the default flow control */
            cbFlow.SelectedIndex = 2;

            comport = new Comport();
            if (comport.isOpen() == false)
                this.Text = "Setup: " + comport.GetPortName() + " port OFF";
            else
                this.Text = "Setup: " + comport.GetPortName() + " port ON";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            /* Todo... */

            if (comport.isOpen() == true)
            {
                comport.Close();
                //MessageBox.Show("COM port failure.");
                //return;
            }

            //comport.init(cbPort.Text, cbBaudrate.Text, cbParity.Text, cbData.Text, cbStop.Text, cbFlow.Text, tbReadTimeOut.Text, tbWriteTimeOut.Text);
            bool result = comport.init(cbPort.Text, Convert.ToInt32(cbBaudrate.Text), cbParity.Text, Convert.ToInt32(cbData.Text),
                cbStop.Text, cbFlow.Text, Convert.ToInt32(tbReadTimeOut.Text), Convert.ToInt32(tbWriteTimeOut.Text));
            //if(result == true) 
            
            comport.Open();
           
            //comport.SetPortDataBits(cbData.Text);
            //comport.SetPortBaudRate(cbBaudrate.Text);
            //comport.SetPortParity(cbParity.Text);
            //comport.SetPortStopBits(cbStop.Text);

            this.Close();
        }












    }
}
