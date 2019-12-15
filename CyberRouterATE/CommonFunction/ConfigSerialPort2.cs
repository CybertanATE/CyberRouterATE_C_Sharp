using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using ComportClass;

namespace CyberRouterATE
{
    public partial class ConfigSerialPort2 : Form
    {
        Comport2 comport2 = null;

        public ConfigSerialPort2()
        {
            InitializeComponent();
        }

        private void ConfigSerialPort2_Load(object sender, EventArgs e)
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

            comport2 = new Comport2();
            if (comport2.isOpen() == false)
                this.Text = "Serial Port 2 -- Setup: " + comport2.GetPortName() + " port OFF";
            else
                this.Text = "Serial Port 2 -- Setup: " + comport2.GetPortName() + " port ON";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            /* Todo... */

            if (comport2.isOpen() == true)
            {
                comport2.Close();
                //MessageBox.Show("COM port failure.");
                //return;
            }

            //comport2.init(cbPort.Text, cbBaudrate.Text, cbParity.Text, cbData.Text, cbStop.Text, cbFlow.Text, tbReadTimeOut.Text, tbWriteTimeOut.Text);
            bool result = comport2.init(cbPort.Text, Convert.ToInt32(cbBaudrate.Text), cbParity.Text, Convert.ToInt32(cbData.Text),
                cbStop.Text, cbFlow.Text, Convert.ToInt32(tbReadTimeOut.Text), Convert.ToInt32(tbWriteTimeOut.Text));
            //if(result == true) 

            comport2.Open();

            //comport2.SetPortDataBits(cbData.Text);
            //comport2.SetPortBaudRate(cbBaudrate.Text);
            //comport2.SetPortParity(cbParity.Text);
            //comport2.SetPortStopBits(cbStop.Text);

            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
