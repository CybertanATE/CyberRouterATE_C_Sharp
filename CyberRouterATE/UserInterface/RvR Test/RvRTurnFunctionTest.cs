///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RvRTurnFunctionTest.cs
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
using System.IO;
using System.Diagnostics;
using System.Threading;
using AgilentInstruments ;
using Ivi.Visa.Interop;
//using NationalInstruments.VisaNS;
using System.Xml;
using System.Net.Sockets;
using System.Net;
using ComportClass;
using System.IO;
using RaspBerryInstruments;
using RouterControlClass;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {     
        //bool bRouterFTThreadRunning = false;        
        string m_RvRTurnTestSubFolder;
        string m_RvRTurnFinalReport;
                
        ///* Declare Test Condition related variable */
        string[,] testConfigRvRTurn; //test condition data        

        Comport comportRvRTurn = null;
        RouterGuiControl guiControl = null;
        turntable turntableRvRTurn = null;

        /* Declare delegate function */
        public delegate void showRvRTurnFunctionTestGUIDelegate();
        public delegate void showRvrTurnTestResultGUI(string[] result);

        private void rvRTurnTableTestToolStripMenuItem_Click(object sender, EventArgs e)
        {        
            sTestItem = "RvR Turn Test";
            ToggleToolStripMenuItem(false);
            tabControl_RouterStartPage.Hide();

            rvRTurnTableTestToolStripMenuItem.Checked = true;
                       
            /* Remove is a bad way
            foreach (TabPage page in RvRTest_tabControl.TabPages)
            {
                if (page.Text.Equals("RvR Test Condition", StringComparison.OrdinalIgnoreCase) || page.Text.Equals("RvR Attenuator Setting", StringComparison.OrdinalIgnoreCase) || page.Text.Equals("RvR-Turn Test Condition", StringComparison.OrdinalIgnoreCase))
                    continue;

                RvRTest_tabControl.TabPages.Remove(page);
            }
             */

            RvRTabHideAllPages();

            tp_RvRTurnTestCondition.Parent = this.RvRTest_tabControl; //Show tabPage
            tp_RvRAttenuatorSetting.Parent = this.RvRTest_tabControl;            
            tp_RvRTurnFunctionTest.Parent = this.RvRTest_tabControl;
            tp_RvRTestResultInfo.Parent = this.RvRTest_tabControl;

            RvRTest_tabControl.Show();

            tsslMessage.Text = RvRTest_tabControl.TabPages[RvRTest_tabControl.SelectedIndex].Text + " Control Panel";            

            string xmlFile = System.Windows.Forms.Application.StartupPath + "\\config\\" + Constants.XML_ROUTER_RvRTURN;
            Debug.WriteLine(File.Exists(xmlFile) ? "File exist" : "File is not existing");

            if (!File.Exists(xmlFile))
            {
                WriteXmlDefaultRvRTurnTest(xmlFile);
            }

            ReadXmlRvRTurnTest(xmlFile);

            string[] _dgvHeader = new string[] { "Angle", "Band", "Mode", "Channel", "Security", "Attenuator", "Tx", "Rx", "Bi" };
            string _condition = "RvR-Turn Test Result Control Panel";

            InitRvRTurnTestCondition();

            InitRouterTestResultTable(_dgvHeader.Length, _dgvHeader, _condition);  
        }
        
        private void chkRvRTurnFunctionTestConfigRouterManually_CheckedChanged(object sender, EventArgs e)
        {
            if (chkRvRTurnFunctionTestConfigRouterManually.Checked)
            {
                labRvRTurnFunctionTestDutModel.Visible = true;
                txtRvRTurnFunctionTestDutModelName.Visible = true;
                cboxRvRTurnFunctionTestModelName.Visible = false;
            }
            else
            {
                labRvRTurnFunctionTestDutModel.Visible = false;
                txtRvRTurnFunctionTestDutModelName.Visible = false;
                cboxRvRTurnFunctionTestModelName.Visible = true;
            }
        }

        private void btnRvRTurnFunctionTestSaveLog_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDebugFileDialog = new SaveFileDialog();

            saveDebugFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            saveDebugFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveDebugFileDialog.FilterIndex = 1;
            saveDebugFileDialog.RestoreDirectory = true;
            saveDebugFileDialog.FileName = "Log";

            if (saveDebugFileDialog.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText(saveDebugFileDialog.FileName, txtRvRTurnFunctionTestInformation.Text);
            }
        }

        private void cboxRvRTurnFunctionTestModelName_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboxRvRTurnFunctionTestModelName.SelectedItem.ToString().IndexOf("TCH") >= 0)
            {
                if (cboxRvRTurnFunctionTestModelName.SelectedIndex == 0) //TCH-2.4G
                {
                    txtRvRTurnFunctionTestDUTIPAddress.Text = "192.168.0.1";
                    txtRvRTurnFunctionTestDutUserName.Text = "";
                    txtRvRTurnFunctionTestDutPassword.Text = "admin";
                }
                if (cboxRvRTurnFunctionTestModelName.SelectedIndex == 1) //TCH-Dual
                {
                    txtRvRTurnFunctionTestDUTIPAddress.Text = "192.168.1.1";
                    txtRvRTurnFunctionTestDutUserName.Text = "admin";
                    txtRvRTurnFunctionTestDutPassword.Text = "password";
                }
            }
            else
            {
                txtRvRTurnFunctionTestDUTIPAddress.Text = "192.168.1.1";
                txtRvRTurnFunctionTestDutUserName.Text = "admin";
                txtRvRTurnFunctionTestDutPassword.Text = "admin";
            }            
        }
 
        private void txtRvRTurnFunctionTestChariotFolder_Click(object sender, EventArgs e)
        {       
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.DefaultExt = "exe";
            openFileDialog1.Title = "Choose Chariot.exe file ";
            openFileDialog1.Filter = "exe files(*.exe)|*.exe|All files(*.*)|*.*";
            openFileDialog1.FileName = "IxChariot.exe";
            openFileDialog1.FilterIndex = 1;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtRvRTurnFunctionTestChariotFolder.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open Chariot.exe: " + ex.Message);
                }
            } 
        }
        
        private void btnRvRTurnFunctionTestRun_Click(object sender, EventArgs e)
        {
            //ForRvRTurnTestOnly();
            //MessageBox.Show("For RvR Turn Test finished!!!");
            //btnRvRTurnFunctionTestRun.Text = "Run";
            //btnRvRTurnFunctionTestRun.Enabled = true;
            //System.Windows.Forms.Cursor.Current = Cursors.Default;
            //return;

           
            /* Prevent double-click from double firing the same action */
            btnRvRTurnFunctionTestRun.Enabled = false;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            btnRvRTurnFunctionTestRun.Text = "Stop" ;

            if (bRouterFTThreadRunning == false)
            {
                bCheckPoint = false;
                if (chkRvRTurnFunctionTestConfigRouterManually.Checked == true) bConfigRouter = false;
                else bConfigRouter = true;

                if (threadRouterFT != null) 
                    threadRouterFT.Abort();

                guiControl = new RouterGuiControl(txtRvRTurnFunctionTestDUTIPAddress.Text);
          
                /* Check if the Dut Control Type and File is ready */
                if (guiControl.checkDutFile())
                {
                    MessageBox.Show("The runtst.exe or the fmttst.exe file doesn't exist, please specify a chariot.exe in which folder!!!", "Warning");
                    btnRvRTurnFunctionTestRun.Text = "Run";
                    btnRvRTurnFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    return;
                }

                /* Check if runtst.exe and fmttst.exe exist */
                string chariotFile = txtRvRTurnFunctionTestChariotFolder.Text;
                string tstFile = chariotFile.Substring(0, chariotFile.Length - Path.GetFileName(chariotFile).Length);
                Path_runtst = tstFile + "runtst.exe";
                Path_fmttst = tstFile + "fmttst.exe";

                SetText("Checking runtst.exe and fmttst.ext whether exist or not.", txtRvRTurnFunctionTestInformation);                
                if (!File.Exists(Path_runtst) || !File.Exists(Path_fmttst))
                {
                    MessageBox.Show("The runtst.exe or the fmttst.exe file doesn't exist, please specify a chariot.exe in which folder!!!","Warning");
                    btnRvRTurnFunctionTestRun.Text = "Run" ;
                    btnRvRTurnFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;                    
                    return;
                }
                SetText("Check result: OK.", txtRvRTurnFunctionTestInformation);

                /* Config and Check Attenuator Setting */
                SetText("Check Test Condition...", txtRvRTurnFunctionTestInformation);
                if (!RvRTurnCheckCondition())
                {
                    MessageBox.Show("Condition check failed. Please check information for detailed!!!", "Warning");
                    btnRvRTurnFunctionTestRun.Text = "Run" ;
                    btnRvRTurnFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;                    
                    return;
                }
                SetText("Check result: OK.", txtRvRTurnFunctionTestInformation);

                /* Check IP address is valid */
                SetText("Check IP address Format.", txtRvRTurnFunctionTestInformation);
                if (!CheckIPValid(txtRvRTurnFunctionTestDUTIPAddress.Text) ||
                   !CheckIPValid(txtRvRTurnFunctionTest24gIPaddress.Text) ||
                   !CheckIPValid(txtRvRTurnFunctionTest5gIPaddress.Text)) 
                {
                    MessageBox.Show("IP address is not valid. Please check your network settings!!!", "Warning");
                    btnRvRTurnFunctionTestRun.Text = "Run" ;
                    btnRvRTurnFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;                     
                    return;
                }
                SetText("Check result: OK.", txtRvRTurnFunctionTestInformation);

                if (chkRvRTurnFunctionTestTurnTable1.Checked)
                {
                    string strRead = string.Empty;
                    /* Create Comport*/
                    comportRvRTurn = new Comport();

                    SetText("Check ComPort.", txtRvRTurnFunctionTestInformation);
                    if (!comportRvRTurn.isOpen())
                    {
                        MessageBox.Show("COM port is not ready! Please check COM port settings!!!", "Warning");
                        btnRvRTurnFunctionTestRun.Text = "Run" ;
                        btnRvRTurnFunctionTestRun.Enabled = true;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;                     
                        return;
                    }
                    SetText("Check result: OK.", txtRvRTurnFunctionTestInformation);

                    SetText("Login to Turn table Raspberry.", txtThroughputFunctionTestInformation);
                    turntableRvRTurn = new turntable(comportRvRTurn);
                    if (!turntableRvRTurn.Login(txtRvRTurnFunctionTestTurntableUsername.Text, txtRvRTurnFunctionTestTurntablePassword.Text, ref strRead))
                    {
                        MessageBox.Show("TurnTable Login Failed!", "Warning");
                        btnRvRTurnFunctionTestRun.Text = "Run";
                        btnRvRTurnFunctionTestRun.Enabled = true;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        return;
                    }

                    SetText(strRead, txtRvRTurnFunctionTestInformation);
                }

                if (bConfigRouter)
                {
                    /* Ping DUT */
                    SetText("Check DUT connection.", txtRvRTurnFunctionTestInformation);
                    if (!PingClient(txtRvRTurnFunctionTestDUTIPAddress.Text, PingTimeout))
                    {
                        MessageBox.Show("Ping DUT failed. Please check your network environment!!!", "Warning");
                        btnRvRTurnFunctionTestRun.Text = "Run";
                        btnRvRTurnFunctionTestRun.Enabled = true;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        return;
                    }
                    SetText("Check result: OK.", txtRvRTurnFunctionTestInformation);
                    model = cboxRvRTurnFunctionTestModelName.Text;

                    security = cboxRvRTurnTestConditionWirelessSecurity.SelectedItem.ToString();
                    passphrase = txtRvRTurnTestConditionPassphrase.Text;
                    if (security.Trim().ToLower() != "none" && passphrase == "")
                    {
                        MessageBox.Show("Setting Passphras for security!!!", "Warning");
                        btnRvRTurnFunctionTestRun.Text = "Run";
                        btnRvRTurnFunctionTestRun.Enabled = true;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        return;
                    }                    
                }
                else //bConfigRouter = false, Manually setting 
                {
                    model = txtRvRTurnFunctionTestDutModelName.Text;
                    if (model == "") model = "RvRTurnTest";
                }
                /* Check if attenuator trun on  */
                SetText("Checking attenuator instrument.", txtRvRTurnFunctionTestInformation);
                if (!CheckAtteuatorOn())
                {
                    MessageBox.Show("Attenuator is not ready. Please turn on the test equipment!!!", "Warning");
                    btnRvRTurnFunctionTestRun.Text = "Run";
                    btnRvRTurnFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    return;
                }
                SetText("Check result: OK.", txtRvRTurnFunctionTestInformation);

                /* Open needed GPIB resource */
                SetText("OpenGPIBPort.", txtRvRTurnFunctionTestInformation);
                OpenGPIBport();                
                                
                /* Read attenuator value for 2.4G and 5G */
                SetText("Read Attenuator Value.", txtRvRTurnFunctionTestInformation);
                ReadAtteunatorValue();

                /* Create sub folder for saving the report data */
                SetText("Create folder.", txtRvRTurnFunctionTestInformation);
                m_RvRTurnTestSubFolder = createRvRTurnTestSubFolder(model);

                dgvRvRTestResultTable.Rows.Clear();

                bRouterFTThreadRunning = true;
                /* Disable all controller */
                ToggleRvRTurnFunctionTestControl(false);

                bCheckPoint = true;

                SetText("============================================", txtRvRTurnFunctionTestInformation);
                SetText("Start RvRTurn testing......", txtRvRTurnFunctionTestInformation);
                                
                threadRouterFT = new Thread(new ThreadStart(DoRvRTurnFunctionTest1));
                threadRouterFT.Name = "RvRTurnTest";
                threadRouterFT.Start();
            }
            else
            {
                if (MessageBox.Show("Do you want to stop the test?", "Stop", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    if (threadRouterFT != null) 
                        threadRouterFT.Abort();
                }
                else
                    return;

                bRouterFTThreadRunning = false;
                btnRvRTurnFunctionTestRun.Text = "Run";
                ToggleRvRTurnFunctionTestControl(true);
            }
            
            /* Button release */
            Thread.Sleep(3000);
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            btnRvRTurnFunctionTestRun.Enabled = true;
        }

        private void DoRvRTurnFunctionTest1()
        {
            string strRead = string.Empty;
            WifiBasic wifi = new WifiBasic();
            AtteuationValue att = new AtteuationValue();
            TstFile tst = new TstFile();
            //string band;
            //string mode;
            //int channel;
            //string ssid;
            //string security;
            //string passphrase;
            //int start;
            //int stop;
            //int steps;
            //string txtst;
            //string rxtst;
            //string bitst;
            string str = string.Empty;            
            Agilent11713A[] a11713 = null;
            decimal[] Attenuation_buttonValue = null;
            string _clientIP = string.Empty;
            int currentAngle = 0;
            int nextAngle = Convert.ToInt32(nudRvRTurnFunctionTestTurnTable1Start.Value);            

            m_RvRTurnFinalReport = System.Windows.Forms.Application.StartupPath + @"\report\" + m_RvRTurnTestSubFolder + @"\finalReport.csv";
            CreateCsvFinalReportRvRTurn(m_RvRTurnFinalReport);
                     
            /* Let the turn table turn to the first angle, angle-start */            
            for (int ang = Convert.ToInt32(nudRvRTurnFunctionTestTurnTable1Start.Value); ang <= Convert.ToInt32(nudRvRTurnFunctionTestTurnTable1Stop.Value); ang += Convert.ToInt32(nudRvRTurnFunctionTestTurnTable1Step.Value))
            {
                if (chkRvRTurnFunctionTestTurnTable1.Checked)
                {  //Run Turn table function  -參數是每次要轉動的角度, 如30度, 而非指定到的角度, 如120度
                    string TurnAngle = (nextAngle - currentAngle).ToString();
                    turntableRvRTurn.ClockWiseTurn(TurnAngle, txtRvRTurnFunctionTestTurntableClockCalibration.Text, ref strRead);
                    currentAngle = nextAngle;
                    nextAngle = currentAngle + Convert.ToInt32(nudRvRTurnFunctionTestTurnTable1Step.Value);
                    str = String.Format("Turn Angle to " + ang.ToString());
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
                    Thread.Sleep(3000);
                }
                else
                {
                    ang = 370;
                }

                /* Set the attenuator to 0 */
                //ConfigureAtteuatorValue(0, a11713, Attenuation_buttonValue);

                /* Main loop for attenuation and configure router */
                for (int i = 0; i < testConfigRvRTurn.GetLength(0); i++)
                {
                    if (bRouterFTThreadRunning == false)
                    {
                        bRouterFTThreadRunning = false;                        
                        MessageBox.Show("Abort test", "Error");
                        this.Invoke(new showRvRTurnFunctionTestGUIDelegate(ToggleRvRTurnFunctionTestGUI));                        
                        threadRouterFT.Abort();                       
                        // Never go here ...
                    }

                    wifi.band = testConfigRvRTurn[i, 0];
                    wifi.mode = testConfigRvRTurn[i, 1];
                    wifi.channel = Int32.Parse(testConfigRvRTurn[i, 2]);
                    wifi.ssid = testConfigRvRTurn[i, 3];
                    wifi.security = testConfigRvRTurn[i, 4];
                    wifi.passphrase = testConfigRvRTurn[i, 5];
                    att.start = Int32.Parse(testConfigRvRTurn[i, 6]);
                    att.stop = Int32.Parse(testConfigRvRTurn[i, 7]);
                    att.steps = Int32.Parse(testConfigRvRTurn[i, 8]);
                    tst.txTst = testConfigRvRTurn[i, 9];
                    tst.rxTst = testConfigRvRTurn[i, 10];
                    tst.biTst = testConfigRvRTurn[i, 11];

                    /* Config Router */
                    str = String.Format("============================================");
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });

                    if (bConfigRouter)
                    {
                        str = String.Format("Configure AP router as modelName:{0}, band:{1}, mode:{2}, channel:{3}, ssid:{4}, security:{5}, Key:{6} ", model, wifi.band, wifi.mode, wifi.channel, wifi.ssid, wifi.security, wifi.passphrase);
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });

                        if (!ConfigureRouter(model, txtRvRTurnFunctionTestDUTIPAddress.Text, txtRvRTurnFunctionTestDutUserName.Text, txtRvRTurnFunctionTestDutPassword.Text, wifi))
                        {
                            str = String.Format("Configure AP router Failed.");
                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
                            continue;
                        }
                        str = String.Format("Configure AP router Succeed.");
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
                        //Thread.Sleep(5000);
                        Thread.Sleep(Convert.ToInt32(nudRvRTurnTestConditionRouterDelayTime.Value) *1000);
                    }

                    if (wifi.band == "2.4G")
                    {
                        a11713 = a11713A_2_4G;
                        Attenuation_buttonValue = Attenuation_buttonValue_2_4G;
                        _clientIP = txtRvRTurnFunctionTest24gIPaddress.Text;
                    }

                    if (wifi.band == "5G")
                    {
                        a11713 = a11713A_5G;
                        Attenuation_buttonValue = Attenuation_buttonValue_5G;
                        _clientIP = txtRvRTurnFunctionTest5gIPaddress.Text;
                    }

                    /* Initial attenuator value that depends on attenuator start */
                    //ConfigureAtteuatorValue(Convert.ToInt32(nudRvRTurnTestConditionAtteuationStart.Value.ToString()), a11713, Attenuation_buttonValue);
                    //ConfigureAtteuatorValue(att.start, a11713, Attenuation_buttonValue);
                    
                    if (RvRTurnMainFunction(model, wifi, att, tst, _clientIP, a11713, Attenuation_buttonValue, ang>360?"X":ang.ToString(), _clientIP) == false)  
                    {  
                        continue;
                    }
                }
            }            

            if (bCheckPoint)
            {
                /* Test finish , write test result to File */
                str = "Test completed!!";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
                
            }
            else
            {
                str = "Some errors was found during testing!!!";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
            }

           
            bRouterFTThreadRunning = false;
            this.Invoke(new showRvRTurnFunctionTestGUIDelegate(ToggleRvRTurnFunctionTestGUI));
        }

        private bool RvRTurnMainFunction(string model, WifiBasic wifi, AtteuationValue att, TstFile tst, string clientIP, Agilent11713A[] a11713, decimal[] Attenuation_buttonValue, string angle, string _clientIP)
        { 
            string str = string.Empty;
            string savePath = createRvRTurnSavePath(m_RvRTurnTestSubFolder, model, wifi, att, angle);
            CreateCsvFileRvRTurn(savePath, model, wifi, att, tst, angle);
            
            string txThroughout = "";
            string rxThroughout = "";
            string biThroughout = "";

            for (int j = att.start; j <= att.stop; j += att.steps)
            {
                if (bRouterFTThreadRunning == false)
                {
                    return false ;
                }

                /* Configure Attenuator value */
                str = "Configure attenuator to [" + j.ToString() + "dB]";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
                ConfigureAtteuatorValue(j, a11713, Attenuation_buttonValue);

                Thread.Sleep(Convert.ToInt32(nudRvRTurnTestConditionAttenuatorDelayTime.Value) * 1000);

                str = "Ping client card : " + _clientIP;
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });

                int nPing = 0;
                while (true)
                {
                    if (PingClient(_clientIP, PingTimeout))
                    {
                        str = "Ping succeed!!!";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
                        break;
                    }

                    if (nPing++ >= Convert.ToInt32(nudRvRTurnFunctionTestPingClientTimeout.Value))
                    {
                        str = "Ping failed!!!";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
                        bCheckPoint = false;
                        break;
                    }
                    Thread.Sleep(1000);
                }

                if (nPing >= Convert.ToInt32(nudRvRTurnFunctionTestPingClientTimeout.Value))
                {
                    txThroughout = "";
                    rxThroughout = "";
                    biThroughout = "";

                    string data1 = String.Format("{0}, {1}, {2}", j.ToString(), txThroughout, rxThroughout, biThroughout);
                    CsvAppend(savePath, data1);                    

                    string[] result1 = new string[] { angle, wifi.band, wifi.mode, wifi.channel.ToString(), wifi.security, j.ToString(), txThroughout, rxThroughout, biThroughout };
                    Invoke(new showRvrTurnTestResultGUI(showRvrTurnTestResult), new object[] { result1 });
                    
                    string finalReport1 = string.Join(",", result1);
                    CsvAppend(m_RvRTurnFinalReport, finalReport1);
                    
                    str = ">>> Next Atteuator value.";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });     

                    continue; //Ping the client card fail.
                }
                
                string outputPath = System.Windows.Forms.Application.StartupPath + @"\report\" + m_RvRTurnTestSubFolder;
                string _inputFile = string.Empty ;
                string _outputHead = @"outpurPATH\Model_Band_Mode_Chanchannel_Security_secMode_Attenuator_attvalue_Angle_angle_";
                string _outputFile = string.Empty;

                /* Replace the _outpurFile with WifiBasic and Attenuation Value*/
                _outputHead = _outputHead.Replace("outpurPATH", outputPath);
                _outputHead = _outputHead.Replace("Model", model);
                _outputHead = _outputHead.Replace("Band", wifi.band);
                _outputHead = _outputHead.Replace("Mode", wifi.mode);
                _outputHead = _outputHead.Replace("channel", wifi.channel.ToString());
                _outputHead = _outputHead.Replace("secMode", wifi.security);
                _outputHead = _outputHead.Replace("attvalue", j.ToString());
                _outputHead = _outputHead.Replace("angle", angle);  

                str = "Start to run Chariot";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });

                /* Run Chariot with Tx.tst */
                if(tst.txTst !="")
                {
                    str = "TX script file: " + tst.txTst;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
                    _inputFile = tst.txTst;
                    _outputFile = _outputHead + "Tx_" + Path.GetFileName(_inputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");
                    RunChariotConsole(Path_runtst, Path_fmttst, _inputFile, _outputFile, txtRvRTurnFunctionTestInformation);
                    string csvFile = _outputFile + ".csv";
                    txThroughout = ThroughputValue(csvFile);
                }

                Thread.Sleep(3000);

                if(tst.rxTst !="")
                {
                    str = "RX script file: " + tst.rxTst;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
                    _inputFile = tst.rxTst;
                    _outputFile = _outputHead + "Rx_" + Path.GetFileName(_inputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");
                    RunChariotConsole(Path_runtst, Path_fmttst, _inputFile, _outputFile, txtRvRTurnFunctionTestInformation);
                    string csvFile = _outputFile + ".csv";
                    rxThroughout = ThroughputValue(csvFile);

                }
                  
                Thread.Sleep(3000);

                if(tst.biTst !="")
                {
                    str = "Bi-Direction script file:" + tst.biTst;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });
                    _inputFile = tst.biTst;
                    _outputFile = _outputHead + "Bi_" + Path.GetFileName(_inputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");
                    RunChariotConsole(Path_runtst, Path_fmttst, _inputFile, _outputFile, txtRvRTurnFunctionTestInformation);
                    string csvFile = _outputFile + ".csv";
                    biThroughout = ThroughputValue(csvFile);

                }

                string data = String.Format("{0}, {1}, {2}", j.ToString(), txThroughout, rxThroughout, biThroughout);
                CsvAppend(savePath, data);                
                
                str = "Chariot Finished.";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });

                string[] result = new string[] { angle, wifi.band, wifi.mode, wifi.channel.ToString(), wifi.security, j.ToString(), txThroughout, rxThroughout, biThroughout };
                Invoke(new showRvrTurnTestResultGUI(showRvrTurnTestResult), new object[] { result });

                string finalReport = string.Join(",", result);
                CsvAppend(m_RvRTurnFinalReport, finalReport);

                txThroughout = "";
                rxThroughout = "";
                biThroughout = "";

                str = ">>> Next Atteuator value.";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRvRTurnFunctionTestInformation });                               
            }

            InserLine2TextFile(savePath, 7, "End Time:, " + DateTime.Now.ToString("yyyy/MM/dd HH:mm")) ;

            return true;
        }
        
        /* Check all needed Condition and Attenuator is ready*/
        private bool RvRTurnCheckCondition()
        {
            /* Check if Test Condition Data is Empty */
            if (dgvRvRTurnTestConditionData.RowCount <= 1)
            {
                MessageBox.Show("Config Test Condition Data Empty, Config First!!!");                
                return false;
            }

            /* Read Test Config and check if TX, RX, Bi tst file exists */
            if (!ReadTestConfig_RvRTurn())
            {
                SetText("Read Test Config Failed.");
                return false ;
            }

            /* Check at least one GPIB number has set */
            if (nudRvRAtteuatorNumber24G.Value == 0 && nudRvRAtteuatorNumber5G.Value == 0)
            {
                MessageBox.Show("Both attenuator number is Zero. Please Check!!!", "Warning");
                return false;
            }

            /* Check if test band and test attenuator is match */
            if (bTest_2_4G)
            {
                if (nudRvRAtteuatorNumber24G.Value == 0)
                {
                    MessageBox.Show("Test Condition 2.4G needed!!!Set 2.4G Attenuator First!!!", "Warning");
                    return false;
                }
            }
            if (bTest_5G)
            {
                if (nudRvRAtteuatorNumber5G.Value == 0)
                {
                    MessageBox.Show("Test Condition 5G needed!!!Set 5G Attenuator First!!!", "Warning");
                    return false;
                }
            }

            /* Check if attenuator number and ip address number is the same */
            if(lboxRvRAtteuationSettingGPIBIP24G.Items.Count != nudRvRAtteuatorNumber24G.Value)
            {
                MessageBox.Show("2.4G: Attenuator number is not equal to the ip number!!!", "Warning");
                return false;
            }

            if(lboxRvRAtteuationSettingGPIBIP5G.Items.Count != nudRvRAtteuatorNumber5G.Value)
            {
                MessageBox.Show("5G: Attenuator number is not equal to the ip number!!!", "Warning");
                return false;
            }            
            return true;
        }

        private bool ReadTestConfig_RvRTurn()
        {
            string[,] rowdata = new string[dgvRvRTurnTestConditionData.RowCount - 1, 12] ;

            bTest_2_4G = false;
            bTest_5G = false;

            for (int i = 0; i < dgvRvRTurnTestConditionData.RowCount-1; i++)
            {
                for (int j = 0; j < 12; j++)
                {
                    rowdata[i, j] = dgvRvRTurnTestConditionData.Rows[i].Cells[j].Value.ToString();
                    /* for j=0 : Set bTest_2_4G=true  if need to test 2.4G , 
                     * bTest_5G = true  if need to test 5G */
                    if (j == 0)
                    {
                        if (rowdata[i, j] == "2.4G" && !bTest_2_4G) bTest_2_4G = true;
                        if (rowdata[i, j] == "5G" && !bTest_5G) bTest_5G = true;
                    }


                    /* For j = 7~9 : Check if tst File exist */
                    if (j >= 9 && j <= 11)
                    {
                        if ( rowdata[i,j]!="" && !File.Exists(rowdata[i, j]))
                        {
                            MessageBox.Show(rowdata[i,j] + " File doesn't exist. Please Check!!", "Warning" );
                            return false;
                        }
                    }
                }                
            }
            testConfigRvRTurn = rowdata;
            return true;
        }        
        
        private string createRvRTurnTestSubFolder(string ModelName)
        {
            string subFolder = ((ModelName == "") ? "E8350_" : ModelName + "_") +
                    DateTime.Now.ToString("yyyyMMdd_HHmmss");


            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);
            
            //if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\LinkRate"))
            //    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\LinkRate");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\PDF"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\PDF");
            
            return subFolder;
        }
        
        private void WriteXmlDefaultRvRTurnTest(string filename)
        {
            XmlWriterSettings setting = new XmlWriterSettings();
            setting.Indent = true; //指定縮排

            XmlWriter writer = XmlWriter.Create(filename, setting);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by RvRTurn Test program.");
            writer.WriteStartElement("CyberRouterATE");
            writer.WriteAttributeString("Item", "RvRTurn Test");

            /* Write Test Condition setting */

            /* Write Attenuator Setting */
            writer.WriteStartElement("Attenuator");

            /* 2.4 G setting */
            writer.WriteStartElement("Band_2_4G");

            /* Write Attenuator Value */
            writer.WriteStartElement("AttenuatorValue_2_4G");
            writer.WriteElementString("X1", "1");
            writer.WriteElementString("X2", "2");
            writer.WriteElementString("X3", "4");
            writer.WriteElementString("X4", "4");
            writer.WriteElementString("X5", "10");
            writer.WriteElementString("X6", "20");
            writer.WriteElementString("X7", "30");
            writer.WriteElementString("X8", "30");
            writer.WriteEndElement();

            /* Write GPIB Interface*/
            writer.WriteStartElement("GPIB_Interface_2_4G");
            writer.WriteElementString("GPIB_Interface", "0");
            writer.WriteEndElement();

            /* Write GPIB Number */
            writer.WriteStartElement("GPIB_Number_2_4G");
            writer.WriteElementString("GPIB_Number", "0");
            writer.WriteEndElement();

            /*Write GPIB IP address*/
            writer.WriteStartElement("GPIB_IP_2_4G");
            
            writer.WriteEndElement();
            /* Writer */


            writer.WriteEndElement();
            //End of write 2.4G

            /* 5G Setting*/
            writer.WriteStartElement("Band_5G");

            /* Write Attenuator Value */
            writer.WriteStartElement("AttenuatorValue_5G");
            writer.WriteElementString("X1", "1");
            writer.WriteElementString("X2", "2");
            writer.WriteElementString("X3", "4");
            writer.WriteElementString("X4", "4");
            writer.WriteElementString("X5", "10");
            writer.WriteElementString("X6", "20");
            writer.WriteElementString("X7", "40");
            writer.WriteElementString("X8", "40");
            writer.WriteEndElement();

            /* Write GPIB Interface*/
            writer.WriteStartElement("GPIB_Interface_5G");
            writer.WriteElementString("GPIB_Interface", "0");
            writer.WriteEndElement();

            /* Write GPIB Number */
            writer.WriteStartElement("GPIB_Number_5G");
            writer.WriteElementString("GPIB_Number", "0");
            writer.WriteEndElement();

            /*Write GPIB IP address*/
            writer.WriteStartElement("GPIB_IP_5G");
            
            writer.WriteEndElement();

            writer.WriteEndElement();
            //End of write 5G

            writer.WriteEndElement();
            //End of write Attenuator


            /* Write Function Test Setting */
            writer.WriteStartElement("FunctionTest");
            // Model section
            writer.WriteStartElement("Model");
            writer.WriteElementString("Name", "0");
            writer.WriteElementString("SN", "123456789");
            writer.WriteElementString("SWVer", "1.0");
            writer.WriteElementString("HWVer", "1.0");
            writer.WriteEndElement();

            // Router Setting
            writer.WriteStartElement("Router");
            writer.WriteElementString("Username", "");
            writer.WriteElementString("Password", "admin");
            writer.WriteElementString("IP", "192.168.0.1");
            writer.WriteEndElement();

            // Client Setting
            writer.WriteStartElement("Client");
            writer.WriteElementString("IPaddr_2_4G", "192.168.1.100");
            writer.WriteElementString("IPaddr_5G", "192.168.1.101");
            writer.WriteElementString("PingTime", "60");
            writer.WriteEndElement();

            //ChariotSetting
            writer.WriteStartElement("Chariot");
            writer.WriteElementString("File", @"C:\Program Files (x86)\Ixia\IxChariot\IxChariot.exe");
            writer.WriteEndElement();            

            //TurnTable 
            writer.WriteStartElement("TurnTable1");
            writer.WriteElementString("Checked", "0");
            writer.WriteElementString("Start", "0");
            writer.WriteElementString("Stop", "360");
            writer.WriteElementString("Step", "90");
            writer.WriteElementString("Username", "pi");
            writer.WriteElementString("Password", "raspberry");
            writer.WriteElementString("ClockwiseCali", "2016");
            writer.WriteElementString("CountClockwiseCali", "2016");
            writer.WriteEndElement();
            writer.WriteEndElement();


            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }

        private void WriteXmlRvRTurnTest(string filename)
        {
            XmlWriterSettings setting = new XmlWriterSettings();
            setting.Indent = true; //指定縮排

            XmlWriter writer = XmlWriter.Create(filename, setting);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by RvRTurn Test program.");
            writer.WriteStartElement("CyberRouterATE");
            writer.WriteAttributeString("Item","RvRTurn Test") ;

            /* Write Test Condition setting */

            /* Write Attenuator Setting */
            writer.WriteStartElement("Attenuator") ;

            /* 2.4 G setting */
            writer.WriteStartElement("Band_2_4G") ;

            /* Write Attenuator Value */
            writer.WriteStartElement("AttenuatorValue_2_4G") ;
            writer.WriteElementString("X1", nudRvRAtteuation24GX1.Value.ToString()) ;
            writer.WriteElementString("X2", nudRvRAtteuation24GX2.Value.ToString()) ;
            writer.WriteElementString("X3", nudRvRAtteuation24GX3.Value.ToString()) ;
            writer.WriteElementString("X4", nudRvRAtteuation24GX4.Value.ToString()) ;
            writer.WriteElementString("X5", nudRvRAtteuation24GX5.Value.ToString()) ;
            writer.WriteElementString("X6", nudRvRAtteuation24GX6.Value.ToString()) ;
            writer.WriteElementString("X7", nudRvRAtteuation24GX7.Value.ToString()) ;
            writer.WriteElementString("X8", nudRvRAtteuation24GX8.Value.ToString()) ;
            writer.WriteEndElement();
            
            /* Write GPIB Interface*/
            writer.WriteStartElement("GPIB_Interface_2_4G");
            writer.WriteElementString("GPIB_Interface", nudRvRGPIBInterface24G.Value.ToString());
            writer.WriteEndElement() ;

            /* Write GPIB Number */
            writer.WriteStartElement("GPIB_Number_2_4G") ;
            writer.WriteElementString("GPIB_Number", nudRvRAtteuatorNumber24G.Value.ToString()) ;
            writer.WriteEndElement();

            /*Write GPIB IP address*/
            writer.WriteStartElement("GPIB_IP_2_4G") ;
            int i=1 ; 
            foreach(string s in lboxRvRAtteuationSettingGPIBIP24G.Items)
            {
                writer.WriteElementString("IP_" + i.ToString(), s) ;
                i++;
            }                
            writer.WriteEndElement();
            /* Writer */


            writer.WriteEndElement();
            //End of write 2.4G

            /* 5G Setting*/
            writer.WriteStartElement("Band_5G") ;

            /* Write Attenuator Value */
            writer.WriteStartElement("AttenuatorValue_5G") ;
            writer.WriteElementString("X1", nudRvRAtteuation5GX1.Value.ToString()) ;
            writer.WriteElementString("X2", nudRvRAtteuation5GX2.Value.ToString()) ;
            writer.WriteElementString("X3", nudRvRAtteuation5GX3.Value.ToString()) ;
            writer.WriteElementString("X4", nudRvRAtteuation5GX4.Value.ToString()) ;
            writer.WriteElementString("X5", nudRvRAtteuation5GX5.Value.ToString()) ;
            writer.WriteElementString("X6", nudRvRAtteuation5GX6.Value.ToString()) ;
            writer.WriteElementString("X7", nudRvRAtteuation5GX7.Value.ToString()) ;
            writer.WriteElementString("X8", nudRvRAtteuation5GX8.Value.ToString()) ;
            writer.WriteEndElement();

            /* Write GPIB Interface*/
            writer.WriteStartElement("GPIB_Interface_5G");
            writer.WriteElementString("GPIB_Interface", nudRvRGPIBInterface5G.Value.ToString());
            writer.WriteEndElement() ;

            /* Write GPIB Number */
            writer.WriteStartElement("GPIB_Number_5G") ;
            writer.WriteElementString("GPIB_Number", nudRvRAtteuatorNumber5G.Value.ToString()) ;
            writer.WriteEndElement();

            /*Write GPIB IP address*/
            writer.WriteStartElement("GPIB_IP_5G") ;
            i = 1;  
            foreach(string s in lboxRvRAtteuationSettingGPIBIP5G.Items)
            {
                writer.WriteElementString("IP_" + i.ToString(), s) ;
                i++;
            }                
            writer.WriteEndElement();

            writer.WriteEndElement();
            //End of write 5G

            writer.WriteEndElement() ;
            //End of write Attenuator


            /* Write Function Test Setting */
            writer.WriteStartElement("FunctionTest");
            // Model section
            writer.WriteStartElement("Model"); 
            writer.WriteElementString("Name", cboxRvRTurnFunctionTestModelName.SelectedIndex.ToString());
            writer.WriteElementString("SN",    txtRvRTurnFunctionTestSerialNumber.Text);
            writer.WriteElementString("SWVer", txtRvRTurnFunctionTestSwVersion.Text);
            writer.WriteElementString("HWVer", txtRvRTurnFunctionTestHwVersion.Text);
            writer.WriteEndElement();

            // Router Setting
            writer.WriteStartElement("Router");
            writer.WriteElementString("Username", txtRvRTurnFunctionTestDutUserName.Text);
            writer.WriteElementString("Password", txtRvRTurnFunctionTestDutPassword.Text);
            writer.WriteElementString("IP", txtRvRTurnFunctionTestDUTIPAddress.Text);                
            writer.WriteEndElement();

            // Client Setting
            writer.WriteStartElement("Client");
            writer.WriteElementString("IPaddr_2_4G",  txtRvRTurnFunctionTest24gIPaddress.Text);
            writer.WriteElementString("IPaddr_5G", txtRvRTurnFunctionTest5gIPaddress.Text);
            writer.WriteElementString("PingTime", nudRvRTurnFunctionTestPingClientTimeout.Value.ToString());  
            writer.WriteEndElement();

            //ChariotSetting
            writer.WriteStartElement("Chariot");
            writer.WriteElementString("File", txtRvRTurnFunctionTestChariotFolder.Text);                
            writer.WriteEndElement();

            //TurnTable 
            writer.WriteStartElement("TurnTable1");
            writer.WriteElementString("Checked", chkRvRTurnFunctionTestTurnTable1.Checked? "1":"0");
            writer.WriteElementString("Start", nudRvRTurnFunctionTestTurnTable1Start.Value.ToString());
            writer.WriteElementString("Stop", nudRvRTurnFunctionTestTurnTable1Stop.Value.ToString());
            writer.WriteElementString("Step", nudRvRTurnFunctionTestTurnTable1Step.Value.ToString());
            writer.WriteElementString("Username", txtRvRTurnFunctionTestTurntableUsername.Text);
            writer.WriteElementString("Password", txtRvRTurnFunctionTestTurntablePassword.Text);
            writer.WriteElementString("ClockwiseCali", txtRvRTurnFunctionTestTurntableClockCalibration.Text);
            writer.WriteElementString("CountClockwiseCali", txtRvRTurnFunctionTestTurntableClockCalibration.Text);
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }

        private bool ReadXmlRvRTurnTest(string filename)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(filename);

            XmlNode node = doc.SelectSingleNode("CyberRouterATE");
            if (node == null)
            {
                return false;
            }

            XmlElement element = (XmlElement)node;
            string strID = element.GetAttribute("Item");
            Debug.Write(strID);
            if (strID.CompareTo("RvRTurn Test") != 0)
            {
                MessageBox.Show("This XML file is incorrect", "Error");
                return false;
            }

            /* Read Attenuator Setting */
            /* Read  2.4G attenuator value x1-x8 */
            XmlNode nodeAttenuator = doc.SelectSingleNode("/CyberRouterATE/Attenuator/Band_2_4G/AttenuatorValue_2_4G");

            try
            {
                string x1 = nodeAttenuator.SelectSingleNode("X1").InnerText;
                string x2 = nodeAttenuator.SelectSingleNode("X2").InnerText;
                string x3 = nodeAttenuator.SelectSingleNode("X3").InnerText;
                string x4 = nodeAttenuator.SelectSingleNode("X4").InnerText;
                string x5 = nodeAttenuator.SelectSingleNode("X5").InnerText;
                string x6 = nodeAttenuator.SelectSingleNode("X6").InnerText;
                string x7 = nodeAttenuator.SelectSingleNode("X7").InnerText;
                string x8 = nodeAttenuator.SelectSingleNode("X8").InnerText;
                Debug.WriteLine("X1: " + x1);
                Debug.WriteLine("X2: " + x2);
                Debug.WriteLine("X3: " + x1);
                Debug.WriteLine("X4: " + x2);
                Debug.WriteLine("X5: " + x1);
                Debug.WriteLine("X6: " + x2);
                Debug.WriteLine("X7: " + x1);
                Debug.WriteLine("X8: " + x2);

                nudRvRAtteuation24GX1.Value = Decimal.Parse(x1);
                nudRvRAtteuation24GX2.Value = Decimal.Parse(x2);
                nudRvRAtteuation24GX3.Value = Decimal.Parse(x3);
                nudRvRAtteuation24GX4.Value = Decimal.Parse(x4);
                nudRvRAtteuation24GX5.Value = Decimal.Parse(x5);
                nudRvRAtteuation24GX6.Value = Decimal.Parse(x6);
                nudRvRAtteuation24GX7.Value = Decimal.Parse(x7);
                nudRvRAtteuation24GX8.Value = Decimal.Parse(x8);
            }
            catch(Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/Attenuator/Bnad_2_4G/Attenuator/" + ex);
            }
            
            /* Read 2.4G GPIB interface value */
            XmlNode nodeInterface_2_4G = doc.SelectSingleNode("/CyberRouterATE/Attenuator/Band_2_4G/GPIB_Interface_2_4G");
            try
            {
                string interface24 = nodeInterface_2_4G.SelectSingleNode("GPIB_Interface").InnerText;
                Debug.WriteLine("GPIB_Interface: " + interface24);
                nudRvRGPIBInterface24G.Value = Decimal.Parse(interface24);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/Attenuator/Band_2_4G/GPIB_Interface_2_4G "+ ex);
            }
            
            /* Read 2.4G GPIB number */
            XmlNode nodeNo_2_4G = doc.SelectSingleNode("/CybertanATE/Attenuator/Band_2_4G/GPIB_Number_2_4G");
            try
            {
                string no24 = nodeNo_2_4G.SelectSingleNode("GPIB_Number").InnerText;
                Debug.WriteLine("GPIB_Number: " + no24);

                nudRvRAtteuatorNumber24G.Value = Decimal.Parse(no24);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/Attenuator/Band_2_4G/GPIB_Number_2_4G " + ex);
            }

            /* Read 2.4G GPIB ip address */
            XmlNode nodeGpibBIP_2_4G = doc.SelectSingleNode("/CyberRouterATE/Attenuator/Band_2_4G/GPIB_IP_2_4G");
            /* Clear all items. Added by James */
            lboxRvRAtteuationSettingGPIBIP24G.Items.Clear();
            int i=1 ;
            for (i = 1; i < nudRvRAtteuatorNumber24G.Value+1; i++)
            {
                string temp = "IP_" + i.ToString();
                string ip = nodeGpibBIP_2_4G.SelectSingleNode("IP_"+i.ToString()).InnerText;
                Debug.WriteLine("IP_"+i.ToString()+" :" +ip);
                lboxRvRAtteuationSettingGPIBIP24G.Items.Add(ip);
            }

            /* Read 5G */
            /* Read 5G attenuator value x1-x8 */
            XmlNode nodeAttenuatorY = doc.SelectSingleNode("/CyberRouterATE/Attenuator/Band_5G/AttenuatorValue_5G");

            try
            {
                string x1 = nodeAttenuatorY.SelectSingleNode("X1").InnerText;
                string x2 = nodeAttenuatorY.SelectSingleNode("X2").InnerText;
                string x3 = nodeAttenuatorY.SelectSingleNode("X3").InnerText;
                string x4 = nodeAttenuatorY.SelectSingleNode("X4").InnerText;
                string x5 = nodeAttenuatorY.SelectSingleNode("X5").InnerText;
                string x6 = nodeAttenuatorY.SelectSingleNode("X6").InnerText;
                string x7 = nodeAttenuatorY.SelectSingleNode("X7").InnerText;
                string x8 = nodeAttenuatorY.SelectSingleNode("X8").InnerText;
                Debug.WriteLine("X1: " + x1);
                Debug.WriteLine("X2: " + x2);
                Debug.WriteLine("X3: " + x1);
                Debug.WriteLine("X4: " + x2);
                Debug.WriteLine("X5: " + x1);
                Debug.WriteLine("X6: " + x2);
                Debug.WriteLine("X7: " + x1);
                Debug.WriteLine("X8: " + x2);

                nudRvRAtteuation5GX1.Value = Decimal.Parse(x1);
                nudRvRAtteuation5GX2.Value = Decimal.Parse(x2);
                nudRvRAtteuation5GX3.Value = Decimal.Parse(x3);
                nudRvRAtteuation5GX4.Value = Decimal.Parse(x4);
                nudRvRAtteuation5GX5.Value = Decimal.Parse(x5);
                nudRvRAtteuation5GX6.Value = Decimal.Parse(x6);
                nudRvRAtteuation5GX7.Value = Decimal.Parse(x7);
                nudRvRAtteuation5GX8.Value = Decimal.Parse(x8);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/Attenuator/Bnad_5G/Attenuator/" + ex);
            }

            /* Read 5G GPIB interface value */
            XmlNode nodeInterface_5G = doc.SelectSingleNode("/CyberRouterATE/Attenuator/Band_5G/GPIB_Interface_5G");
            try
            {
                string interface5G = nodeInterface_5G.SelectSingleNode("GPIB_Interface").InnerText;
                Debug.WriteLine("GPIB_Interface: " + interface5G);
                nudRvRGPIBInterface5G.Value = Decimal.Parse(interface5G);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/Attenuator/Band_5G/GPIB_Interface_5G " + ex);
            }

            /* Read 5G GPIB number */
            XmlNode nodeNo_5G = doc.SelectSingleNode("/CyberRouterATE/Attenuator/Band_5G/GPIB_Number_5G");
            try
            {
                string no5G = nodeNo_5G.SelectSingleNode("GPIB_Number").InnerText;
                Debug.WriteLine("GPIB_Number: " + no5G);

                nudRvRAtteuatorNumber5G.Value = Decimal.Parse(no5G);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/Attenuator/Band_5G/GPIB_Number_5G " + ex);
            }

            /* Read 2.4G GPIB ip address */
            XmlNode nodeGpibBIP_5G = doc.SelectSingleNode("/CyberRouterATE/Attenuator/Band_5G/GPIB_IP_5G");
            /* Clear all items. Added by James */
            lboxRvRAtteuationSettingGPIBIP5G.Items.Clear() ;
            try
            {
                for (i = 1; i < nudRvRAtteuatorNumber5G.Value + 1; i++)
                {
                    string temp = "IP_" + i.ToString();
                    string ip = nodeGpibBIP_5G.SelectSingleNode("IP_" + i.ToString()).InnerText;
                    Debug.WriteLine("IP_" + i.ToString() + " :" + ip);
                    lboxRvRAtteuationSettingGPIBIP5G.Items.Add(ip);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/Attenuator/Band_5G/GPIB_IP_5G " + ex);
            }            

            /* Read Function Test Setting */
            /* Read Model */
            XmlNode nodeModel = doc.SelectSingleNode("/CyberRouterATE/FunctionTest/Model");
            try
            {
                string Name = nodeModel.SelectSingleNode("Name").InnerText;
                string SN = nodeModel.SelectSingleNode("SN").InnerText;
                string SWVer = nodeModel.SelectSingleNode("SWVer").InnerText;
                string HWVer = nodeModel.SelectSingleNode("HWVer").InnerText;

                cboxRvRTurnFunctionTestModelName.SelectedIndex = Int32.Parse(Name);
                Debug.WriteLine("Name: " + Name + ": " + cboxRvRTurnFunctionTestModelName.SelectedItem.ToString());
                Debug.WriteLine("SN: " + SN);
                Debug.WriteLine("SWVer: " + SWVer);
                Debug.WriteLine("HWVer: " + HWVer);    
                
                txtRvRTurnFunctionTestSerialNumber.Text = SN;
                txtRvRTurnFunctionTestSwVersion.Text = SWVer;
                txtRvRTurnFunctionTestHwVersion.Text = HWVer;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/FunctionTest/Model " + ex);
            }

            /* Read Router Setting */
            XmlNode nodeRouter = doc.SelectSingleNode("/CyberRouterATE/FunctionTest/Router");
            try
            {
                string Username = nodeRouter.SelectSingleNode("Username").InnerText;
                string Password = nodeRouter.SelectSingleNode("Password").InnerText;
                string DutIP = nodeRouter.SelectSingleNode("IP").InnerText;

                Debug.WriteLine("Username: " + Username);
                Debug.WriteLine("Password: " + Password) ;
                Debug.WriteLine("DUT IP: " + DutIP);

                txtRvRTurnFunctionTestDutUserName.Text = Username;
                txtRvRTurnFunctionTestDutPassword.Text = Password;
                txtRvRTurnFunctionTestDUTIPAddress.Text = DutIP;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/FunctionTest/Router " + ex);
            }

            /* Read Client Setting */
            XmlNode nodeClient = doc.SelectSingleNode("/CyberRouterATE/FunctionTest/Client");
            try
            {
                string IP_24G = nodeClient.SelectSingleNode("IPaddr_2_4G").InnerText;
                string IP_5G = nodeClient.SelectSingleNode("IPaddr_5G").InnerText;
                string Pingtime = nodeClient.SelectSingleNode("PingTime").InnerText;

                Debug.WriteLine("IP_2_4G: "+IP_24G);
                Debug.WriteLine("IP_5G: "+IP_5G);
                Debug.WriteLine("PingTime: " + Pingtime);

                txtRvRTurnFunctionTest24gIPaddress.Text = IP_24G;
                txtRvRTurnFunctionTest5gIPaddress.Text = IP_5G;
                nudRvRTurnFunctionTestPingClientTimeout.Value = Convert.ToDecimal(Pingtime);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/FunctionTest/Client " + ex);
            }
        
            /* Read Chariot Folder */
            XmlNode nodeChariot = doc.SelectSingleNode("/CyberRouterATE/FunctionTest/Chariot");
            try
            {
                string cFile = nodeChariot.SelectSingleNode("File").InnerText;

                Debug.WriteLine("File: "+ cFile);

                txtRvRTurnFunctionTestChariotFolder.Text = cFile;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/FunctionTest/Chariot " + ex);
            }

            /* Read TurnTable1 */
            XmlNode nodeTurnTable1 = doc.SelectSingleNode("/CyberRouterATE/FunctionTest/TurnTable1");
            try
            {
                string Check = nodeTurnTable1.SelectSingleNode("Checked").InnerText;
                string Start = nodeTurnTable1.SelectSingleNode("Start").InnerText;
                string Stop = nodeTurnTable1.SelectSingleNode("Stop").InnerText;
                string Step = nodeTurnTable1.SelectSingleNode("Step").InnerText;
                string Username = nodeTurnTable1.SelectSingleNode("Username").InnerText;
                string Password = nodeTurnTable1.SelectSingleNode("Password").InnerText;
                string ClockwiseCali = nodeTurnTable1.SelectSingleNode("ClockwiseCali").InnerText;
                //string CountClockwiseCali = nodeTurnTable1.SelectSingleNode("CountClockwiseCali").InnerText;
                                
                Debug.WriteLine("Checked: " + Check);
                Debug.WriteLine("Start: " + Start);
                Debug.WriteLine("Stop: " + Stop);
                Debug.WriteLine("Step: " + Step);
                Debug.WriteLine("Username: " + Username);
                Debug.WriteLine("Password: " + Password);
                Debug.WriteLine("ClockwiseCali: " + Step);
                //Debug.WriteLine("CountClockwiseCali: " + CountClockwiseCali);
                
                chkRvRTurnFunctionTestTurnTable1.Checked = (Check =="1"? true:false);
                nudRvRTurnFunctionTestTurnTable1Start.Value =Convert.ToDecimal(Start);
                nudRvRTurnFunctionTestTurnTable1Stop.Value =Convert.ToDecimal(Stop);
                nudRvRTurnFunctionTestTurnTable1Step.Value =Convert.ToDecimal(Step);                
                txtRvRTurnFunctionTestTurntableUsername.Text = Username;
                txtRvRTurnFunctionTestTurntablePassword.Text = Password;
                txtRvRTurnFunctionTestTurntableClockCalibration.Text = ClockwiseCali;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/FunctionTest/TurnTable1 " + ex);
            }       
            
            return true;
        }
        
        private void ToggleRvRTurnFunctionTestGUI()
        {
            ToggleRvRTurnFunctionTestControl(true);
            Debug.WriteLine("Toggle");
        }

        private void ToggleRvRTurnFunctionTestControl(bool Toggle)
        {
            /* Test Condition */
            rbtnRvRTurnTestCondition24G.Enabled = Toggle;
            rbtnRvRTurnTestCondition5G.Enabled = Toggle;
            cboxRvRTurnTestConditionWirelessMode.Enabled = Toggle;
            cboxRvRTurnTestConditionWirelessChannel.Enabled = Toggle;
            txtRvRTurnTestConditionSsid.Enabled = Toggle;
            cboxRvRTurnTestConditionWirelessSecurity.Enabled = Toggle;
            txtRvRTurnTestConditionPassphrase.Enabled = Toggle;
            nudRvRTurnTestConditionRouterDelayTime.Enabled = Toggle;
            
            nudRvRTurnTestConditionAtteuationStart.Enabled = Toggle;
            nudRvRTurnTestConditionAtteuationStop.Enabled = Toggle;
            nudRvRTurnTestConditionAtteuationStep.Enabled = Toggle;
            nudRvRTurnTestConditionAttenuatorDelayTime.Enabled = Toggle;

            txtRvRTurnTestConditionTxTst.Enabled = Toggle;
            txtRvRTurnTestConditionRxTst.Enabled = Toggle;
            txtRvRTurnTestConditionBiTst.Enabled = Toggle;
            btnRvRTurnTestConditionAddSetting.Enabled = Toggle;
            btnRvRTurnTestConditionEditSetting.Enabled = Toggle;
            btnRvRTurnTestConditionSaveSetting.Enabled = Toggle;
            btnRvRTurnTestConditionLoadSetting.Enabled = Toggle;            

            /* Attenuator Setting */
            nudRvRAtteuation24GX1.Enabled = Toggle;
            nudRvRAtteuation24GX2.Enabled = Toggle;
            nudRvRAtteuation24GX3.Enabled = Toggle;
            nudRvRAtteuation24GX4.Enabled = Toggle;
            nudRvRAtteuation24GX5.Enabled = Toggle;
            nudRvRAtteuation24GX6.Enabled = Toggle;
            nudRvRAtteuation24GX7.Enabled = Toggle;
            nudRvRAtteuation24GX8.Enabled = Toggle;
            nudRvRGPIBInterface24G.Enabled = Toggle;
            nudRvRAtteuatorNumber24G.Enabled = Toggle;
            nudRvRAtteuationGPIBIPAdderss24G.Enabled = Toggle;
            labRvRAttenuationAdd24G.Enabled = Toggle;
            labRvRAttenuationDel24G.Enabled = Toggle;
            lboxRvRAtteuationSettingGPIBIP24G.Enabled = Toggle;

            nudRvRAtteuation5GX1.Enabled = Toggle;
            nudRvRAtteuation5GX2.Enabled = Toggle;
            nudRvRAtteuation5GX3.Enabled = Toggle;
            nudRvRAtteuation5GX4.Enabled = Toggle;
            nudRvRAtteuation5GX5.Enabled = Toggle;
            nudRvRAtteuation5GX6.Enabled = Toggle;
            nudRvRAtteuation5GX7.Enabled = Toggle;
            nudRvRAtteuation5GX8.Enabled = Toggle;
            nudRvRGPIBInterface5G.Enabled = Toggle;
            nudRvRAtteuatorNumber5G.Enabled = Toggle;
            nudRvRAttenuationGPIBIPAddress5G.Enabled = Toggle;
            labRvRAttenuationAdd5G.Enabled = Toggle;
            labRvRAttenuationDel5G.Enabled = Toggle;
            lboxRvRAtteuationSettingGPIBIP5G.Enabled = Toggle;

            btn_AttenuationSetting_GPIB_Info.Enabled = Toggle;

            /* Function Test */
            cboxRvRTurnFunctionTestModelName.Enabled = Toggle;
            txtRvRTurnFunctionTestSerialNumber.Enabled = Toggle;
            txtRvRTurnFunctionTestSwVersion.Enabled = Toggle;
            txtRvRTurnFunctionTestHwVersion.Enabled = Toggle;

            chkRvRTurnFunctionTestConfigRouterManually.Enabled = Toggle;
            txtRvRTurnFunctionTestDutModelName.Enabled = Toggle;
                        
            txtRvRTurnFunctionTestDutUserName.Enabled = Toggle;
            txtRvRTurnFunctionTestDutPassword.Enabled = Toggle;

            txtRvRTurnFunctionTestDUTIPAddress.Enabled = Toggle;
            txtRvRTurnFunctionTest24gIPaddress.Enabled = Toggle;
            txtRvRTurnFunctionTest5gIPaddress.Enabled = Toggle;
            txtRvRTurnFunctionTestChariotFolder.Enabled = Toggle;
            nudRvRTurnFunctionTestPingClientTimeout.Enabled = Toggle;

            chkRvRTurnFunctionTestTurnTable1.Enabled = Toggle;
            nudRvRTurnFunctionTestTurnTable1Start.Enabled = Toggle;
            nudRvRTurnFunctionTestTurnTable1Stop.Enabled = Toggle;
            nudRvRTurnFunctionTestTurnTable1Step.Enabled = Toggle;
            txtRvRTurnFunctionTestTurntableUsername.Enabled = Toggle;
            txtRvRTurnFunctionTestTurntablePassword.Enabled = Toggle;
            txtRvRTurnFunctionTestTurntableClockCalibration.Enabled = Toggle;
            btnRvRTurnFunctionTestSaveLog.Enabled = Toggle;

            btnRvRTurnFunctionTestRun.Text = Toggle ? "Run" : "Stop";
            Debug.WriteLine(btnRvRTurnFunctionTestRun.Text);

            /* Disable/Enable "Items" and "Setup" menu item */
            itemToolStripMenuItem.Enabled = Toggle;
            setupToolStripMenuItem.Enabled = Toggle;

            /* Show total elasped testing time */
            if (!Toggle)
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                swElapsedTime.Restart();
                timerElaspedTime.Enabled = true;
                timerElaspedTime.Start();
            }
            else
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;

                swElapsedTime.Stop();
                timerElaspedTime.Stop();
                timerElaspedTime.Enabled = false;
            }

            if (bRouterFTThreadRunning == false)
            {
                MessageBox.Show(this, "Test complete!!!", "Information", MessageBoxButtons.OK);
            }            
        }

        private void showRvrTurnTestResult(string[] result)
        {
            dgvRvRTestResultTable.Rows.Add(result);
        }
              
        private string createRvRTurnSavePath(string subFolder, string model, WifiBasic wifi, AtteuationValue att, string angle)
        {
            string PathFile = @"\report\SUBFOLDER\Report_MODEL_BAND_MODE_CHAN_channel_Security_secMode_Attenuator_start_stop_steps_Angel_angle_DATE.csv";
            
            PathFile = PathFile.Replace("SUBFOLDER", subFolder);
            PathFile = PathFile.Replace("MODEL", model);
            PathFile = PathFile.Replace("BAND", wifi.band);
            PathFile = PathFile.Replace("MODE", wifi.mode);
            PathFile = PathFile.Replace("channel", wifi.channel.ToString());
            PathFile = PathFile.Replace("secMode", wifi.security);
            PathFile = PathFile.Replace("start", att.start.ToString());
            PathFile = PathFile.Replace("stop", att.stop.ToString());
            PathFile = PathFile.Replace("steps",att.steps.ToString());
            PathFile = PathFile.Replace("angle", angle);
            PathFile = PathFile.Replace("DATE", DateTime.Now.ToString("yyyyMMdd-HHmm"));          
            
            PathFile = System.Windows.Forms.Application.StartupPath + PathFile;

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

            return PathFile;
        }

        private bool CreateCsvFileRvRTurn(string filePath, string model, WifiBasic wifi, AtteuationValue att, TstFile tst, string angle)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    FileStream myfile = File.Create(filePath);
                    Thread.Sleep(1000);
                    myfile.Close();
                }

                StringBuilder csv = new StringBuilder();

                //string str = "Start Auto Run Testing Time : " + DateTime.Now.DayOfWeek.ToString()  +DateTime.Now.ToString("yyyy-MM-dd HH-mm");
                /* Blank */
                string str = "";
                csv.AppendLine(str);

                str = "CyberATE RvR + Turn Table Test Report"; //Title
                csv.AppendLine(str);

                /* Blank */
                str = "";
                csv.AppendLine(str);


                str = String.Format("{0}, {1}", "Test", "CyberATE " + model + " RvR + Turn Table Test Report"); //Test Name
                csv.AppendLine(str);

                str = String.Format("{0}, {1}", "Station", "CyberTAN ATE-" + model); 
                csv.AppendLine(str);

                str = String.Format("{0}, {1}", "Start Time", DateTime.Now.ToString("yyyy/MM/dd HH:mm")); 
                csv.AppendLine(str);

                //str = String.Format("{0}, {1}", "End Time", ""); 
                //csv.AppendLine(str);

                str = String.Format("{0}, {1}", "Model", model); 
                csv.AppendLine(str);

                str = String.Format("Wireless Setting:,  Band:{0}   Mode:{1}  SSID:{2}   Channel:{3}   Security:{4} ", wifi.band, wifi.mode, wifi.ssid, wifi.channel, wifi.security);
                csv.AppendLine(str);

                str = String.Format("Atteunator:,  start:{0}  stop:{1}   step:{2} ", att.start, att.stop, att.steps);
                csv.AppendLine(str);

                str = String.Format("Angle:,  {0} ", angle);
                csv.AppendLine(str);

                str = String.Format("Tx Tst File:, {0} ", tst.txTst);
                csv.AppendLine(str);

                str = String.Format("Rx Tst File:, {0} ", tst.rxTst);
                csv.AppendLine(str);

                str = String.Format("Bi Tst File:, {0} ", tst.biTst);
                csv.AppendLine(str);

                str = "";
                csv.AppendLine(str);

                str = String.Format("{0}, {1}, {2}, {3}", "Attenuator", "Tx", "Rx", "Bi" );
                csv.AppendLine(str);
               
                //File.WriteAllText(filePath, csv.ToString()); 
                File.AppendAllText(filePath, csv.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Create csv file " + filePath + "Error: " + ex.ToString());
                return false;
            }

            return true;
        }

        private bool CreateCsvFinalReportRvRTurn(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    FileStream myfile = File.Create(filePath);
                    Thread.Sleep(1000);
                    myfile.Close();
                }

                StringBuilder csv = new StringBuilder();

                //string str = "Start Auto Run Testing Time : " + DateTime.Now.DayOfWeek.ToString()  +DateTime.Now.ToString("yyyy-MM-dd HH-mm");
                /* Blank */
                string str = "";
                csv.AppendLine(str);

                str = "CyberATE RvR + Turn Table Final Report"; //Title
                csv.AppendLine(str);

                /* Blank */
                str = "";
                csv.AppendLine(str);

                string[] _dgvHeader = new string[] { "Angle", "Band", "Mode", "Channel", "Security", "Attenuator", "Tx", "Rx", "Bi" };
                str = string.Join(",",_dgvHeader);
                csv.AppendLine(str);
                //File.WriteAllText(filePath, csv.ToString()); 
                File.AppendAllText(filePath, csv.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Create Final report " + filePath + "Error: " + ex.ToString());
                return false;
            }          

            return true;
        }
       
        private void ForRvRTurnTestOnly()
        {
            bRouterFTThreadRunning = true;
            ToggleRvRTurnFunctionTestControl(false);
            for (int dim = 1; dim <= 100; dim++)
            {
                Thread.Sleep(3000);
            }

            //int d = 0;
        }

    } //End of class 
}
