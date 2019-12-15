///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : ThroughputTestFunctionTest.cs
///  Update         : 2017-03-30
///  Description    : All Channel Throughput Test Condtion 
///  Modified       : 2017-03-30 Initial version  
///
///  Comments       : 
///---------------------------------------------------------------------------------------
///

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
using System.Xml;
using System.Net.Sockets;
using System.Net;
using ComportClass;
using RaspBerryInstruments;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {  
        bool bThroughputFunctionTestRunning = false;
        string m_ThroughputTestSubFolder;        
        string m_ThroughputFinalReport;

        //string Path_runtst;
        //string Path_fmttst;
        Thread threadThroughputFunctionTest;
        Comport comportThroughput = null;
        turntable turntableThroughput = null;
        string modelType = string.Empty;

                
        /* Declare Test Condition related variable */
        string[,] testConfigThroughput;
        //bool bTest_2_4G = false;
        //bool bTest_5G = false;
        //int PingTimeout = 1000;
        //string model;
        //string security = "None";
        //string passphrase = string.Empty;
        //bool bConfigRouter = true;

        //bool bCheckPoint;
        //string WiFiData;
        //StringBuilder linkrate_csv;
        
        //public delegate void showThroughputFunctionTestGUIDelegate();
        
        private void throughputTestToolStripMenuItem_Click(object sender, EventArgs e)
        {                      
            sTestItem = TestItemConstants.TESTITEM_THROUGHPUT;            
            ToggleToolStripMenuItem(false);
            tabControl_RouterStartPage.Hide();

            throughputTestToolStripMenuItem.Checked = true;       

            tabControl_Throughput.Show();
            
            tsslMessage.Text = tabControl_Throughput.TabPages[tabControl_Throughput.SelectedIndex].Text + " Control Panel";

            string xmlFile = System.Windows.Forms.Application.StartupPath + "\\config\\" + Constants.XML_ROUTER_THROUGHPUT;
            Debug.WriteLine(File.Exists(xmlFile) ? "File exist" : "File is not existing");

            if (!File.Exists(xmlFile))
            {
                WriteXmlDefaultThroughputTest(xmlFile);
            }

            ReadXmlThroughputTest(xmlFile);
            InitThroughputTestCondition();           
        }

        private void btnThroughputFunctionTestChariotFolder_Click(object sender, EventArgs e)
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
                    txtThroughputFunctionTestChariotFolder.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open Chariot.exe: " + ex.Message);
                }
            } 
        }

        private void btnThroughputFunctionTestSaveLog_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDebugFileDialog = new SaveFileDialog();

            saveDebugFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            saveDebugFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveDebugFileDialog.FilterIndex = 1;
            saveDebugFileDialog.RestoreDirectory = true;
            saveDebugFileDialog.FileName = "Log";

            if (saveDebugFileDialog.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText(saveDebugFileDialog.FileName, txtThroughputFunctionTestInformation.Text);
            }
        }
        
        private void btnThroughputFunctionTestRun_Click(object sender, EventArgs e)       
        {
            WiFiData = "";
            linkrate_csv = new StringBuilder();
            /* Prevent double-click from double firing the same action */
            btnThroughputFunctionTestRun.Enabled = false;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            btnThroughputFunctionTestRun.Text = "Stop" ;            

            if (bThroughputFunctionTestRunning == false)
            {
                bCheckPoint = false;
                if (chkThroughputFunctionTestConfigRouterManually.Checked == true)
                    bConfigRouter = false;
                else bConfigRouter = true;
                
                if (threadThroughputFunctionTest != null) 
                    threadThroughputFunctionTest.Abort(); 
                              
                /* Check if runtst.exe and fmttst.exe exist */
                string chariotFile = txtThroughputFunctionTestChariotFolder.Text;
                string tstFile = chariotFile.Substring(0, chariotFile.Length - Path.GetFileName(chariotFile).Length);
                Path_runtst = tstFile + "runtst.exe";
                Path_fmttst = tstFile + "fmttst.exe";

                SetText("Checking runtst.exe and fmttst.ext whether exist or not.", txtThroughputFunctionTestInformation);
                if (!File.Exists(Path_runtst) || !File.Exists(Path_fmttst))
                {
                    MessageBox.Show("The runtst.exe or the fmttst.exe file doesn't exist, please specify a chariot.exe in which folder!!!","Warning");
                    btnThroughputFunctionTestRun.Text = "Run" ;
                    btnThroughputFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;                    
                    return;
                }
                SetText("Check result: OK.", txtThroughputFunctionTestInformation);
                
                /* Check and read test condition */
                SetText("Check Test Condition...", txtThroughputFunctionTestInformation);
                //SetText("Checking the attenuator settings whether a correct value.");
                if (!ThroughputTestCheckCondition())
                {
                    MessageBox.Show("Condition check failed. Please Check!!!", "Warning");
                    btnThroughputFunctionTestRun.Text = "Run" ;
                    btnThroughputFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;                    
                    return;
                }
                SetText("Check result: OK.", txtThroughputFunctionTestInformation);

                /* Check IP address is valid */
                SetText("Check IP address Format.", txtThroughputFunctionTestInformation);
                if (!CheckIPValid(txtThroughputFunctionTestDUTIPAddress.Text) ||
                   !CheckIPValid(txtThroughputFunctionTest24gIPaddress.Text) ||
                   !CheckIPValid(txtThroughputFunctionTest5gIPaddress.Text)) 
                {
                    MessageBox.Show("IP address is not valid. Please check your network settings!!!", "Warning");
                    btnThroughputFunctionTestRun.Text = "Run" ;
                    btnThroughputFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;                     
                    return;
                }
                SetText("Check result: OK.", txtThroughputFunctionTestInformation);

                if (chkThroughputFunctionTestTurnTable1.Checked)
                {
                    string strRead = string.Empty;
                    /* Create Comport*/
                    comportThroughput = new Comport();

                    SetText("Check ComPort.", txtThroughputFunctionTestInformation);
                    if (!comportThroughput.isOpen())
                    {
                        MessageBox.Show("COM port is not ready! Please check COM port settings!!!", "Warning");
                        btnThroughputFunctionTestRun.Text = "Run";
                        btnThroughputFunctionTestRun.Enabled = true;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        return;
                    }
                    SetText("Check result: OK.", txtThroughputFunctionTestInformation);

                    SetText("Login to Turn table Raspberry.", txtThroughputFunctionTestInformation);
                    turntableThroughput = new turntable(comportThroughput);
                    if (!turntableThroughput.Login(txtThroughputFunctionTestTurntableUsername.Text, txtThroughputFunctionTestTurntablePassword.Text, ref strRead))
                    {
                        MessageBox.Show("TurnTable Login Failed!", "Warning");
                        btnThroughputFunctionTestRun.Text = "Run";
                        btnThroughputFunctionTestRun.Enabled = true;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        return;
                    }

                    SetText(strRead, txtThroughputFunctionTestInformation);
                }

                if (bConfigRouter)
                {
                    /* Ping DUT */
                    SetText("Check DUT connection.", txtThroughputFunctionTestInformation);
                    if (!PingClient(txtThroughputFunctionTestDUTIPAddress.Text, PingTimeout))
                    {
                        MessageBox.Show("Ping DUT failed. Please check your network environment!!!", "Warning");
                        btnThroughputFunctionTestRun.Text = "Run";
                        btnThroughputFunctionTestRun.Enabled = true;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        return;
                    }
                    SetText("Check result: OK.", txtThroughputFunctionTestInformation);

                    modelType = cboxThroughputFunctionTestModelName.Text;                   
                }

                //else //bConfigRouter = false, Manually setting 
                //{
                //    model = txtThroughputFunctionTestDutModelName.Text;
                //}

                model = txtThroughputFunctionTestDutModelName.Text;

                                
                /* Create sub folder for saving the report data */
                m_ThroughputTestSubFolder = createThroughputTestSubFolder(model);                
                bThroughputFunctionTestRunning = true;
                /* Disable all controller */
                ToggleThroughputFunctionTestControl(false);

                bCheckPoint = true;                
                
                SetText("============================================");
                SetText("Start Throughput testing.......");
                                
                threadThroughputFunctionTest = new Thread(new ThreadStart(DoThroughputFunctionTest));
                threadThroughputFunctionTest.Name = "ThroughputTest";
                threadThroughputFunctionTest.Start();
            }
            else
            {
                if (MessageBox.Show("Do you want to stop the test?", "Stop", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    if (threadThroughputFunctionTest != null) 
                        threadThroughputFunctionTest.Abort();
                }
                else
                    return;

                bThroughputFunctionTestRunning = false;
                btnThroughputFunctionTestRun.Text = "Run";
                ToggleThroughputFunctionTestControl(true);
            }
            
            /* Button release */
            Thread.Sleep(3000);
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            btnThroughputFunctionTestRun.Enabled = true;
        }

        private void DoThroughputFunctionTest()
        {
            WirelessParameter wpParameter = new WirelessParameter();
            TstFile tstFile = new TstFile();

            //WirelessParameter wpParameter = new WirelessParameter();
            //TstFile tstFile = new TstFile();

            //string band;
            //string mode;
            //string ssid_config;
            //string ssid_text;
            //string channel_config;
            //int channel_start;
            //int channel_stop;
            //string security_config;
            //string security_mode;
            //string key_index;
            //string passphrase;
            //string txtst;
            //string rxtst;
            //string bitst;
            string strRead = string.Empty;
            string str = string.Empty;
            string currentTime;
            int currentAngle = 0;
            int nextAngle = Convert.ToInt32(nudThroughputFunctionTestTurnTable1Start.Value);

            m_ThroughputFinalReport = System.Windows.Forms.Application.StartupPath + @"\report\" + m_ThroughputTestSubFolder + @"\finalReport.csv";
            CreateCsvFinalReportThroughput(m_ThroughputFinalReport);

            /* Let the turn table turn to the first angle, angle-start */
            for (int ang = Convert.ToInt32(nudThroughputFunctionTestTurnTable1Start.Value); ang <= Convert.ToInt32(nudThroughputFunctionTestTurnTable1Stop.Value); ang += Convert.ToInt32(nudThroughputFunctionTestTurnTable1Step.Value))
            {
                if (chkThroughputFunctionTestTurnTable1.Checked)
                {  //Run Turn table function  -參數是每次要轉動的角度, 如30度, 而非指定到的角度, 如120度
                    string TurnAngle = (nextAngle - currentAngle).ToString();
                    turntableThroughput.ClockWiseTurn(TurnAngle, txtThroughputFunctionTestTurntableClockCalibration.Text, ref strRead);
                    currentAngle = nextAngle;
                    nextAngle = currentAngle + Convert.ToInt32(nudThroughputFunctionTestTurnTable1Step.Value);
                    str = String.Format("Turn Angle to " + ang.ToString());
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtThroughputFunctionTestInformation });
                    Thread.Sleep(3000);
                }
                else
                {
                    ang = 370;
                }

                /* Main loop for configuring router */
                for (int i = 0; i < testConfigThroughput.GetLength(0); i++)
                {
                    if (bThroughputFunctionTestRunning == false)
                    {
                        bThroughputFunctionTestRunning = false;
                        MessageBox.Show("Abort test", "Error");
                        this.Invoke(new showRouterGUIDelegate(ToggleThroughputFunctionTestGUI));
                        threadThroughputFunctionTest.Abort();
                        // Never go here ...
                    }

                    wpParameter.band = testConfig[i, 0];
                    wpParameter.mode = testConfig[i, 1];
                    wpParameter.ssid_config = testConfig[i, 2];
                    wpParameter.ssid_text = testConfig[i, 3];
                    wpParameter.channel_config = testConfig[i, 4];
                    try
                    {
                        wpParameter.channel_start = Int32.Parse(testConfig[i, 5]);
                        wpParameter.channel_stop = Int32.Parse(testConfig[i, 6]);
                    }
                    catch (Exception)
                    {
                        Debug.WriteLine("Channel nnumber is not an interger");
                        bThroughputFunctionTestRunning = false;
                        MessageBox.Show("Channel nnumber is not an interger", "Error");
                        this.Invoke(new showRouterGUIDelegate(ToggleThroughputFunctionTestGUI));
                        threadThroughputFunctionTest.Abort();
                        // Never go here ...
                    }

                    wpParameter.security_config = testConfig[i, 7];
                    wpParameter.security_mode = testConfig[i, 8];
                    wpParameter.key_index = testConfig[i, 9];
                    wpParameter.passphrase = testConfig[i, 10];
                    tstFile.txTst = testConfig[i, 11];
                    tstFile.rxTst = testConfig[i, 12];
                    tstFile.biTst = testConfig[i, 13];

                    /* Config Router */
                    str = String.Format("===============    Test Condition ===================");
                    Invoke(new SetTextCallBack(SetText), new object[] { str });
                    Invoke(new showWirelessPatameterContentDelegate(showWirelessPatameterPartContent), new object[] { wpParameter, txtThroughputFunctionTestInformation });

                    for (int c = wpParameter.channel_start; c <= wpParameter.channel_stop; c++)
                    {
                        if (bThroughputFunctionTestRunning == false)
                        {
                            bThroughputFunctionTestRunning = false;
                            MessageBox.Show("Abort test", "Error");
                            this.Invoke(new showRouterGUIDelegate(ToggleThroughputFunctionTestGUI));
                            threadThroughputFunctionTest.Abort();
                            // Never go here ...
                        }

                        str = String.Format("Configure AP router, channel: {0}", c);
                        Invoke(new SetTextCallBack(SetText), new object[] { str, txtThroughputFunctionTestInformation });

                        if (!ConfigureRouterX5(modelType, wpParameter, c))
                        {
                            str = String.Format("Configure AP router Failed:");
                            Invoke(new SetTextCallBack(SetText), new object[] { str, txtThroughputFunctionTestInformation });
                            Invoke(new showWirelessPatameterContentDelegate(showWirelessPatameterPartContent), new object[] { wpParameter, txtThroughputFunctionTestInformation });

                            continue;
                        }

                        str = String.Format("Configure AP router Finished");
                        Invoke(new SetTextCallBack(SetText), new object[] { str, txtThroughputFunctionTestInformation });










                    }
















                    //        for (int c = channelstart; c <= channelend; c++)
                    //        {
                    //            if (bThroughputFunctionTestRunning == false)
                    //            {
                    //                break;
                    //            }

                    //            str = String.Format("Configure AP router as modelName:{0}, band:{1}, mode:{2}, ssid:{3}, channel:{4} ", model, band, mode, ssid, c);
                    //            Invoke(new SetTextCallBack(SetText), new object[] { str });

                    //            if (!ConfigureRouter(model, band, mode, c, ssid, security, passphrase))
                    //                continue;
                    //            str = String.Format("Configure AP router Succeed.");
                    //            Invoke(new SetTextCallBack(SetText), new object[] { str });

                    //            str = "Ping 2.4G client IP and checking endpoint whether be activated...";
                    //            Invoke(new SetTextCallBack(SetText), new object[] { str });

                    //            int nPing = 0;
                    //            while (true)
                    //            {
                    //                if (PingClient(this.txt_ClientSetting_24gIPaddress.Text, PingTimeout))
                    //                {
                    //                    str = "Ping succeed!!!";
                    //                    Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //                    break;
                    //                }

                    //                if (nPing++ == 60)
                    //                {
                    //                    str = "Ping failed!!!";
                    //                    Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //                    bCheckPoint = false;
                    //                    break;
                    //                }
                    //                Thread.Sleep(1000);
                    //            }


                    //            string input = testConfig[i, 5];
                    //            string outputPath = System.Windows.Forms.Application.StartupPath + @"\report\" + m_ThroughputTestSubFolder;
                    //            /* file name is E8350_2_4G_20M_channel_1_attenuatot_1*/
                    //            /* E8350_TX_2_4G_11n_20M_Channel_" + i.ToString() + "dB_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss") */
                    //            string outputHeader = @"\" + model + "_2_4G_11N_" + mode + "_Channel_" + c.ToString() + "_";
                    //            string output = string.Empty;

                    //            currentTime = DateTime.Now.ToString("yyyy_MMdd_HHmmss");

                    //            str = "Start to run Chariot";
                    //            Invoke(new SetTextCallBack(SetText), new object[] { str });

                    //            string path = System.Windows.Forms.Application.StartupPath + @"\report\" + m_ThroughputTestSubFolder + @"\LinkRate";
                    //            string resultFile = path + "\\LinkResult.csv";

                    //            Thread.Sleep(10000);
                    //            /* Run Tx */
                    //            if (testConfig[i, 7] != "")
                    //            {
                    //                str = "TX script file: " + testConfig[i, 7];
                    //                Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //                input = testConfig[i, 7];
                    //                output = outputPath + outputHeader + "Tx_" + Path.GetFileName(input) + "_" + currentTime;

                    //                RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                    //                // To record remote site's wifi data
                    //                //WiFiData = GetWiFiData();
                    //                //string linkRate = model + "_2_4G_11N_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Tx_" + Path.GetFileName(input) + "_" + currentTime;
                    //                //WriteResultCSV(linkRate + ".csv", resultFile);
                    //            }
                    //            //Thread.Sleep(15000);
                    //            Thread.Sleep(10000);
                    //            /* Run RX */
                    //            if (testConfig[i, 8] != "")
                    //            {
                    //                str = "RX script file: " + testConfig[i, 8];
                    //                Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //                input = testConfig[i, 8];
                    //                output = outputPath + outputHeader + "Rx_" + Path.GetFileName(input) + "_" + currentTime;

                    //                RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                    //                // To record remote site's wifi data
                    //                //WiFiData = GetWiFiData();
                    //                //string linkRate = model + "_2_4G_11N_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Rx_" + Path.GetFileName(input) + "_" + currentTime;
                    //                //WriteResultCSV(linkRate + ".csv", resultFile);
                    //            }
                    //            //Thread.Sleep(15000);
                    //            Thread.Sleep(10000);
                    //            /* Run bi-Direction */
                    //            if (testConfig[i, 9] != "")
                    //            {
                    //                str = "Bi-Direction script file: " + testConfig[i, 9];
                    //                Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //                input = testConfig[i, 9];
                    //                output = outputPath + outputHeader + "Bi_" + Path.GetFileName(input) + "_" + currentTime;

                    //                RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                    //                //// To record remote site's wifi data
                    //                //WiFiData = GetWiFiData();
                    //                //string linkRate = model + "_2_4G_11N_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Bi_" + Path.GetFileName(input) + "_" + currentTime;
                    //                //WriteResultCSV(linkRate + ".csv", resultFile);
                    //            }

                    //            str = "Chariot Finished.";
                    //            Invoke(new SetTextCallBack(SetText), new object[] { str }); 
                    //        }                   
                    //    }

                    //    if(band == "5G")
                    //    {   
                    //        int indexStart = -1 ;
                    //        int indexEnd = -1 ;
                    //        for(int tmp = 0; tmp < Channel_5G_All.Length; tmp++)
                    //        {
                    //            if(Channel_5G_All[tmp] == channelstart)
                    //            {
                    //                indexStart = tmp ;
                    //            }

                    //            if(Channel_5G_All[tmp] == channelend)
                    //            {
                    //                indexEnd = tmp ;
                    //            }

                    //        }

                    //        if(indexStart == -1 || indexEnd == -1)
                    //        {
                    //            continue;
                    //        }


                    //        for (int c = indexStart; c <= indexEnd; c++)
                    //        {
                    //            if (bThroughputFunctionTestRunning == false)
                    //            {
                    //                break;
                    //            }

                    //            str = String.Format("Configure AP router as modelName:{0}, band:{1}, mode:{2}, ssid:{3}, channel:{4} ", model, band, mode, ssid, Channel_5G_All[c]);
                    //            Invoke(new SetTextCallBack(SetText), new object[] { str });

                    //            if (!ConfigureRouter(model, band, mode, Channel_5G_All[c], ssid, security, passphrase))
                    //                continue;
                    //            str = String.Format("Configure AP router Succeed.");
                    //            Invoke(new SetTextCallBack(SetText), new object[] { str });

                    //            str = "Ping 5G client IP and checking endpoint whether be activated...";
                    //            Invoke(new SetTextCallBack(SetText), new object[] { str });

                    //            int nPing = 0;
                    //            while(true)
                    //            {
                    //                if (PingClient(this.txt_ClientSetting_5gIPaddress.Text, PingTimeout))
                    //                {
                    //                    str = "Ping succeed!!!";
                    //                    Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //                    break;
                    //                }

                    //                if (nPing++ == 60)
                    //                {
                    //                    str = "Ping failed!!!";
                    //                    Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //                    bCheckPoint = false;
                    //                    break;
                    //                }
                    //                Thread.Sleep(1000);
                    //            }                        

                    //            string input = testConfig[i, 5];
                    //            string outputPath = System.Windows.Forms.Application.StartupPath + @"\report\" + m_ThroughputTestSubFolder;
                    //            /* file name is E8350_2_4G_20M_channel_1_attenuatot_1*/
                    //            /* E8350_TX_2_4G_11n_20M_Channel_" + i.ToString() + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss") */
                    //            string outputHeader = @"\" + model + "_5G_" + mode + "_Channel_" + Channel_5G_All[c].ToString() + "_";
                    //            string output = string.Empty;                        

                    //            currentTime = DateTime.Now.ToString("yyyy_MMdd_HHmmss");

                    //            str = "Start to run Chariot";
                    //            Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //            //RunChariotConsole(Path_runtst, Path_fmttst, input, output);                        

                    //            string path = System.Windows.Forms.Application.StartupPath + @"\report\" + m_ThroughputTestSubFolder + @"\LinkRate";
                    //            string resultFile = path + "\\LinkResult.csv";

                    //            Thread.Sleep(10000);
                    //            /* Run Tx */
                    //            if (testConfig[i, 7] != "")
                    //            {
                    //                str = "TX script file: " + testConfig[i, 7];
                    //                Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //                input = testConfig[i, 7];
                    //                output = outputPath + outputHeader + "Tx_" + Path.GetFileName(input) + "_" + currentTime;

                    //                RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                    //                // To record remote site's wifi data
                    //                //WiFiData = GetWiFiData();
                    //                //string linkRate = model + "_5G_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Tx_" + Path.GetFileName(input) + "_" + currentTime;
                    //                //WriteResultCSV(linkRate + ".csv", resultFile);
                    //            }
                    //            //Thread.Sleep(15000);
                    //            Thread.Sleep(10000);
                    //            /* Run RX */
                    //            if (testConfig[i, 8] != "")
                    //            {
                    //                str = "RX script file: " + testConfig[i, 8];
                    //                Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //                input = testConfig[i, 8];
                    //                output = outputPath + outputHeader + "Rx_" + Path.GetFileName(input) + "_" + currentTime;

                    //                RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                    //                // To record remote site's wifi data
                    //                //WiFiData = GetWiFiData();
                    //                //string linkRate = model + "_5G_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Rx_" + Path.GetFileName(input) + "_" + currentTime;
                    //                //WriteResultCSV(linkRate + ".csv", resultFile);
                    //            }
                    //            //Thread.Sleep(15000);
                    //            Thread.Sleep(10000);
                    //            /* Run Bi-Direction */
                    //            if (testConfig[i, 9] != "")
                    //            {
                    //                str = "Bi-Direction script file: " + testConfig[i, 9];
                    //                Invoke(new SetTextCallBack(SetText), new object[] { str });
                    //                input = testConfig[i, 9];
                    //                output = outputPath + outputHeader + "Bi_" + Path.GetFileName(input) + "_" + currentTime;

                    //                RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                    //                // To record remote site's wifi data
                    //                //WiFiData = GetWiFiData();
                    //                //string linkRate = model + "_5G_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Bi_" + Path.GetFileName(input) + "_" + currentTime;
                    //                //WriteResultCSV(linkRate + ".csv", resultFile);
                    //            }

                    //            str = "Chariot Finished.";
                    //            Invoke(new SetTextCallBack(SetText), new object[] { str });                        
                    //        }                    
                    //    }
                    //}



                    if (bCheckPoint)
                    {
                        /* Test finish , write test result to File */
                        str = "Test completed!!";
                        Invoke(new SetTextCallBack(SetText), new object[] { str });

                    }
                    else
                    {
                        str = "Some errors was found during testing!!!";
                        Invoke(new SetTextCallBack(SetText), new object[] { str });
                    }



                    /* Save data to result.csv */
                    str = "Save data to result.csv file.";
                    Invoke(new SetTextCallBack(SetText), new object[] { str });
                    string pathcsv = System.Windows.Forms.Application.StartupPath + @"\report\" + m_ThroughputTestSubFolder;
                    string resultCsv = pathcsv + "\\Result.csv";

                    if (File.Exists(resultCsv))
                        File.Delete(resultCsv);

                    string[] files = Directory.GetFiles(pathcsv);
                    WriteCSV(files, resultCsv);

                    bThroughputFunctionTestRunning = false;
                    //this.Invoke(new showThroughputFunctionTestGUIDelegate(ToggleThroughputFunctionTestGUI));
                }
            }
        }

        ///* Check all needed Condition and Attenuator is ready*/
        //private bool CheckCondition()
        //{
        //    /* Check if Test Condition Data is Empty */
        //    if (dataGridView1.RowCount <= 1)
        //    {
        //        MessageBox.Show("Config Test Condition Data Empty, Config First!!!");                
        //        return false;
        //    }

        //    /* Read Test Config and check if TX, RX, Bi tst file exists */
        //    if (!ReadTestConfig())
        //    {
        //        SetText("Read Test Config Failed.");
        //        return false ;
        //    }            
        //    return true;
        //}        
       
        //private bool ReadTestConfig()
        //{
        //    string[,] rowdata = new string[dataGridView1.RowCount - 1, 10] ;

        //    bTest_2_4G = false;
        //    bTest_5G = false;

        //    for (int i = 0; i < dataGridView1.RowCount-1; i++)
        //    {
        //        for (int j = 0; j < 10; j++)
        //        {
        //            rowdata[i, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
        //            /* for j=0 : Set bTest_2_4G=true  if need to test 2.4G , 
        //             * bTest_5G = true  if need to test 5G */
        //            if (j == 0)
        //            {
        //                if (rowdata[i, j] == "2.4G" && !bTest_2_4G) bTest_2_4G = true;
        //                if (rowdata[i, j] == "5G" && !bTest_5G) bTest_5G = true;
        //            }


        //            /* For j = 7~9 : Check if tst File exist */
        //            if (j >= 7 && j <= 9)
        //            {
        //                if ( rowdata[i,j]!="" && !File.Exists(rowdata[i, j]))
        //                {
        //                    MessageBox.Show(rowdata[i,j] + " File doesn't exist. Please Check!!", "Warning" );
        //                    return false;
        //                }
        //            }
        //        }                
        //    }
        //    testConfig = rowdata;
        //    return true;
        //}

        private void ToggleFunctionTestControl(bool Toggle)
        {
            /* Test Condition */
            rbtnThroughputTestCondition24G.Enabled = Toggle;
            rbtnThroughputTestCondition5G.Enabled = Toggle;
            cboxThroughputTestConditionWirelessMode.Enabled = Toggle;
            cbox_Model_Name.Enabled = Toggle;
            //cboxThroughputTestConditionAllChannel.Enabled = Toggle;
            nudThroughputTestConditionWirelessChannelStart.Enabled = Toggle;
            //nudThroughputTestConditionWirelessChannelStop.Enabled = Toggle;
            txtThroughputTestConditionSSIDText.Enabled = Toggle;
            cboxThroughputTestConditionSecurity.Enabled = Toggle;
            txtThroughputTestConditionPassphrase.Enabled = Toggle;
            txtThroughputTestConditionTxTst.Enabled = Toggle;
            txtThroughputTestConditionRxTst.Enabled = Toggle;
            txtThroughputTestConditionBiTst.Enabled = Toggle;
            btnThroughputTestConditionAddSetting.Enabled = Toggle;
            btnThroughputTestConditionEditSetting.Enabled = Toggle;




           
            /* Function Test */
            cbox_Model_Name.Enabled = Toggle;
            txt_Model_SerialNumber.Enabled = Toggle;
            txt_Model_HWVersion.Enabled = Toggle;
            txt_Model_SWVersion.Enabled = Toggle;
            
            /* Commented by James - for EA8500 */
            //txt_RouterSetting_UserName.Enabled = Toggle;
            //txt_RouterSetting_Password.Enabled = Toggle;
            
            
            txt_RouterSetting_DUTIPAddress.Enabled = Toggle;
            txt_ClientSetting_24gIPaddress.Enabled = Toggle;
            txt_ClientSetting_5gIPaddress.Enabled = Toggle;
            txt_ChariotSetting_ChariotFolder.Enabled = Toggle;

            btnThroughputFunctionTestRun.Text = Toggle ? "Run" : "Stop";
            Debug.WriteLine(btnThroughputFunctionTestRun.Text);

            if (bThroughputFunctionTestRunning == false)
            {
                MessageBox.Show(this, "Test complete!!!", "Information", MessageBoxButtons.OK);
            }
        }

        private string createThroughputTestSubFolder(string ModelName)
        {
            string subFolder = ((ModelName == "") ? "E8350_" : ModelName + "_") +
                    DateTime.Now.ToString("yyyy_MMdd_HHmmss");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);
            
            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\LinkRate"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\LinkRate");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\PDF"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\PDF");
            
            return subFolder;
        }          

        //private int[] ArrayStringToInt(string[] str)
        //{
        //    int[] temp = new int[str.Length];
        //    int i =0 ;
        //    foreach (string s in str)
        //    {
        //        temp[i++] = Int32.Parse(s);
        //    }
        //    return temp;
        //}

        private void WriteXmlDefaultThroughputTest(string filename)
        {
            XmlWriterSettings setting = new XmlWriterSettings();
            setting.Indent = true; //指定縮排

            XmlWriter writer = XmlWriter.Create(filename, setting);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by Throughput Test program.");
            writer.WriteStartElement("CybertanATE");
            writer.WriteAttributeString("Item", "Throughput Test");

            /* Write Test Condition setting */

            /* Write Function Test Setting */
            writer.WriteStartElement("FunctionTest");
            // Model section
            writer.WriteStartElement("Model");
            writer.WriteElementString("Name", cbox_Model_Name.SelectedIndex.ToString());
            writer.WriteElementString("SN", txt_Model_SerialNumber.Text);
            writer.WriteElementString("SWVer", txt_Model_SWVersion.Text);
            writer.WriteElementString("HWVer", txt_Model_HWVersion.Text);
            writer.WriteEndElement();

            // Router Setting
            writer.WriteStartElement("Router");
            writer.WriteElementString("Username", txt_RouterSetting_UserName.Text);
            writer.WriteElementString("Password", txt_RouterSetting_Password.Text);
            writer.WriteElementString("IP", txt_RouterSetting_DUTIPAddress.Text);
            writer.WriteEndElement();

            // Client Setting
            writer.WriteStartElement("Client");
            writer.WriteElementString("IPaddr_2_4G", txt_ClientSetting_24gIPaddress.Text);
            writer.WriteElementString("IPaddr_5G", txt_ClientSetting_5gIPaddress.Text);
            writer.WriteEndElement();

            //ChariotSetting
            writer.WriteStartElement("Chariot");
            writer.WriteElementString("File", txt_ChariotSetting_ChariotFolder.Text);
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }

        private void WriteXmlThroughputTest(string filename)
        {
            XmlWriterSettings setting = new XmlWriterSettings();
            setting.Indent = true; //指定縮排

            XmlWriter writer = XmlWriter.Create(filename, setting);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by Throughput Test program.");
            writer.WriteStartElement("CybertanATE") ;
            writer.WriteAttributeString("Item","Throughput Test") ;

            /* Write Test Condition setting */
            
            /* Write Function Test Setting */
            writer.WriteStartElement("FunctionTest");
            // Model section
            writer.WriteStartElement("Model");
            writer.WriteElementString("Name", cbox_Model_Name.SelectedIndex.ToString());
            writer.WriteElementString("SN", txt_Model_SerialNumber.Text);
            writer.WriteElementString("SWVer", txt_Model_SWVersion.Text);
            writer.WriteElementString("HWVer", txt_Model_HWVersion.Text);
            writer.WriteEndElement();

            // Router Setting
            writer.WriteStartElement("Router");
            writer.WriteElementString("Username", txt_RouterSetting_UserName.Text);
            writer.WriteElementString("Password", txt_RouterSetting_Password.Text);
            writer.WriteElementString("IP", txt_RouterSetting_DUTIPAddress.Text);                
            writer.WriteEndElement();

            // Client Setting
            writer.WriteStartElement("Client");
            writer.WriteElementString("IPaddr_2_4G", txt_ClientSetting_24gIPaddress.Text);
            writer.WriteElementString("IPaddr_5G", txt_ClientSetting_5gIPaddress.Text);                
            writer.WriteEndElement();

            //ChariotSetting
            writer.WriteStartElement("Chariot");
            writer.WriteElementString("File", txt_ChariotSetting_ChariotFolder.Text);                
            writer.WriteEndElement();            
            writer.WriteEndElement();

            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }

        private bool ReadXmlThroughputTest(string filename)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(filename);

            XmlNode node = doc.SelectSingleNode("CybertanATE");
            if (node == null)
            {
                return false;
            }

            XmlElement element = (XmlElement)node;
            string strID = element.GetAttribute("Item");
            Debug.Write(strID);
            if (strID.CompareTo("Throughput Test") != 0)
            {
                MessageBox.Show("This XML file is incorrect", "Error");
                return false;
            }            

            /* Read Function Test Setting */
            /* Read Model */
            XmlNode nodeModel = doc.SelectSingleNode("/CybertanATE/FunctionTest/Model");
            try
            {
                string Name = nodeModel.SelectSingleNode("Name").InnerText;
                string SN = nodeModel.SelectSingleNode("SN").InnerText;
                string SWVer = nodeModel.SelectSingleNode("SWVer").InnerText;
                string HWVer = nodeModel.SelectSingleNode("HWVer").InnerText;

                cbox_Model_Name.SelectedIndex = Int32.Parse(Name);
                Debug.WriteLine("Name: " + Name + ": " + cbox_Model_Name.SelectedItem.ToString());
                Debug.WriteLine("SN: " + SN);
                Debug.WriteLine("SWVer: " + SWVer);
                Debug.WriteLine("HWVer: " + HWVer);    
                
                txt_Model_SerialNumber.Text = SN;
                txt_Model_SWVersion.Text = SWVer;
                txt_Model_HWVersion.Text = HWVer;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/FunctionTest/Model " + ex);
            }

            /* Read Router Setting */
            XmlNode nodeRouter = doc.SelectSingleNode("/CybertanATE/FunctionTest/Router");
            try
            {
                string Username = nodeRouter.SelectSingleNode("Username").InnerText;
                string Password = nodeRouter.SelectSingleNode("Password").InnerText;
                string DutIP = nodeRouter.SelectSingleNode("IP").InnerText;

                Debug.WriteLine("Username: " + Username);
                Debug.WriteLine("Password: " + Password) ;
                Debug.WriteLine("DUT IP: " + DutIP);

                txt_RouterSetting_UserName.Text = Username;
                txt_RouterSetting_Password.Text = Password;
                txt_RouterSetting_DUTIPAddress.Text = DutIP;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/FunctionTest/Router " + ex);
            }

            /* Read Client Setting */
            XmlNode nodeClient = doc.SelectSingleNode("/CybertanATE/FunctionTest/Client");
            try
            {
                string IP_24G = nodeClient.SelectSingleNode("IPaddr_2_4G").InnerText;
                string IP_5G = nodeClient.SelectSingleNode("IPaddr_5G").InnerText;

                Debug.WriteLine("IP_2_4G: "+IP_24G);
                Debug.WriteLine("IP_5G: "+IP_5G);

                txt_ClientSetting_24gIPaddress.Text = IP_24G;
                txt_ClientSetting_5gIPaddress.Text = IP_5G;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/FunctionTest/Client " + ex);
            }
        
            /* Read Chariot Folder */
            XmlNode nodeChariot = doc.SelectSingleNode("/CybertanATE/FunctionTest/Chariot");
            try
            {
                string cFile = nodeChariot.SelectSingleNode("File").InnerText;

                Debug.WriteLine("File: "+ cFile);

                txt_ChariotSetting_ChariotFolder.Text = cFile;

            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/FunctionTest/Chariot " + ex);
            }
            
            return true;
        }

        //private bool ConfigureRouter(string modelName,string band, string mode, int channel, string ssid, string security, string key)
        //{
        //    string cgi= string.Empty;
        //    //string CgiHeader = "http://" + txt_RouterSetting_DUTIPAddress.Text + "/mfgtst.cgi?"; 
           
        //    switch (modelName)
        //    {
        //        case "E8350":
        //                cgi = "http://" + txt_RouterSetting_DUTIPAddress.Text + "/mfgtst.cgi?" + "sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=1";
        //                if(band == "2.4G")
        //                {
        //                    string iMode = mode.Substring(0, mode.Length - 1);
        //                    if (channel == 0)
        //                    { /* channel auto */
        //                        cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=Auto");
        //                    }
        //                    else
        //                    {
        //                        if (iMode == "40")
        //                        {  /* when mode is 40M, need to switch channel to channel -2 */


        //                            if ((channel - 2) >= 0)
        //                                cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + (channel - 2).ToString());
        //                            else
        //                                cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=-1");
        //                        }
        //                        if (iMode == "20")
        //                        {
        //                            cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + channel.ToString());
        //                        }
        //                    }

        //                    cgi = cgi.Replace("sys_BandWidth=20", "sys_BandWidth=" + iMode);
        //                    cgi = cgi.Replace("sys_wlSSID=ssid", "sys_wlSSID=" + ssid);
        //                }

        //                /* cgi => sys_wlInterface=0&sys_wlMode=an-mixed&sys_BandWidth=20&sys_wlSSID=HWSQA5G&sys_wlChannel=1 */
        //                if (band == "5G")
        //                {
        //                    string[] s = mode.Split('_');
        //                    if (s[0] == "11N") cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=an-mixed");
        //                    if (s[0] == "11AC") cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=ac-mixed");
        //                    string iMode = s[1].Substring(0, s[1].Length - 1);

        //                    cgi = cgi.Replace("sys_wlInterface=0", "sys_wlInterface=1");
        //                    cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=an-mixed");
        //                    cgi = cgi.Replace("sys_BandWidth=20", "sys_BandWidth="+iMode );
        //                    cgi = cgi.Replace("sys_wlSSID=ssid", "sys_wlSSID=" + ssid);
        //                    if (channel == 0)
        //                    { /* channel auto */
        //                        cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=Auto");
        //                    }
        //                    else
        //                    {
        //                        cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + channel.ToString());
        //                    }
        //                }

        //                if (!SendCGICommand(cgi, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text, "200 OK"))
        //                {
        //                    return false;
        //                }
                        
        //                Thread.Sleep(5000);
                    
        //            break;
        //        case "EA8500X5":
        //            string host = txt_RouterSetting_DUTIPAddress.Text;

        //            /* configure router with telnet connection */
        //            TelnetConnection tc = new TelnetConnection(host, 23);
        //            string str = tc.Login("", "", 100);
        //            Invoke(new SetTextCallBack(SetText), new object[] { str });
        //            string prompt = str.TrimEnd();

        //            //while connected
        //            if (tc.IsConnected)
        //            {
        //                /* display server output */
        //                str = tc.Read();
        //                Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                /* send client input to server */
        //                prompt = Console.ReadLine();
        //                tc.WriteLine(prompt);

        //                /* display server outpur */
        //                str = tc.Read();
        //                Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                //tc.WriteLine("ls");
        //                //str = tc.Read();
        //                //Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                /* Configure router band, channel and ssid */

        //                str = " wifi_test.sh ap ";
        //                if (band == "2.4G")
        //                {
        //                    str += "2g ";
        //                    str += channel.ToString() + " " + ssid;
        //                    tc.WriteLine(str);
        //                    while (true)
        //                    {

        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                        if (str.IndexOf("#") != -1) break;
        //                    }


        //                    /* Configure mode */
        //                    string smode = mode.Substring(0, 2);
        //                    if (smode != "40")
        //                    {
        //                        //str = "uci set wireless.@wifi-device[0].hwmode=11ng";
        //                        str = "uci set wireless.@wifi-iface[0].hwmode=11ng";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(1000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                        //str = "uci set wireless.@wifi-iface[0].htmode=VHT" + smode;
        //                        str = "uci set wireless.@wifi-device[0].htmode=HT" + smode;
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(1000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                        /* Configure security */

        //                        /* Reset wifi */
        //                        str = "wifi down";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                        str = "wifi up";
        //                        tc.WriteLine(str);
        //                        while (true)
        //                        {

        //                            Thread.Sleep(3000);
        //                            str = tc.Read();
        //                            if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                            if (str.IndexOf("#") != -1) break;
        //                        }
        //                        //Invoke(new SetTextCallBack(SetText), new object[] { str });


        //                        str = "iwpriv wifi0 setCountryID 843";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });



        //                    }

        //                    if (smode == "40")
        //                    {
        //                        //str = "uci set wireless.@wifi-device[0].hwmode=11ng";
        //                        str = "uci set wireless.@wifi-iface[0].hwmode=11ng";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(1000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                        str = "uci set wireless.@wifi-iface[0].htmode=VHT" + smode;
        //                        //str = "uci set wireless.@wifi-device[0].htmode=HT" + smode;
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(1000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                        /* Configure security */

        //                        /* Reset wifi */
        //                        str = "wifi down";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                        str = "wifi up";
        //                        tc.WriteLine(str);
        //                        while (true)
        //                        {

        //                            Thread.Sleep(3000);
        //                            str = tc.Read();
        //                            if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                            if (str.IndexOf("#") != -1) break;
        //                        }
        //                        //Invoke(new SetTextCallBack(SetText), new object[] { str });


        //                        str = "iwpriv wifi0 setCountryID 843";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                    }


        //                    /* Show wireless status */
        //                    str = "uci show wireless";
        //                    tc.WriteLine(str);
        //                    Thread.Sleep(3000);
        //                    str = tc.Read();
        //                    Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                    /* Show speed */
        //                    str = "iwconfig";
        //                    tc.WriteLine(str);
        //                    Thread.Sleep(3000);
        //                    str = tc.Read();
        //                    Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                }

        //                if (band == "5G")
        //                {
        //                    str += "5g ";
        //                    str += channel.ToString() + " " + ssid;
        //                    tc.WriteLine(str);
        //                    while (true)
        //                    {

        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                        if (str.IndexOf("#") != -1) break;
        //                    }



        //                    /* Configure mode */
        //                    string[] s = mode.Split('_');
        //                    string smode = s[1].Substring(0, 2);

        //                    if (smode == "20")
        //                    {
        //                        if (s[0] == "11N")
        //                        {
        //                            str = "uci set wireless.@wifi-iface[1].hwmode=11an";
        //                        }
        //                        if (s[1] == "11AC")
        //                        {
        //                            str = "uci set wireless.@wifi-iface[1].hwmode=11ac";
        //                        }

        //                        tc.WriteLine(str);
        //                        Thread.Sleep(1000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });


        //                        //str = "uci set wireless.@wifi-iface[1].htmode=VHT" + smode;
        //                        str = "uci set wireless.@wifi-device[1].htmode=HT" + smode;
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(1000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });


        //                        /* Configure security */

        //                        /* Reset wifi */
        //                        str = "wifi down";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                        str = "wifi up";
        //                        tc.WriteLine(str);
        //                        while (true)
        //                        {

        //                            Thread.Sleep(3000);
        //                            str = tc.Read();
        //                            if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                            if (str.IndexOf("#") != -1) break;
        //                        }

        //                        str = "iwpriv wifi0 setCountryID 843";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                    }

        //                    if (smode == "40")
        //                    {
        //                        if (s[0] == "11N")
        //                        {
        //                            str = "uci set wireless.@wifi-iface[1].hwmode=11an";
        //                        }
        //                        if (s[1] == "11AC")
        //                        {
        //                            str = "uci set wireless.@wifi-iface[1].hwmode=11ac";
        //                        }

        //                        tc.WriteLine(str);
        //                        Thread.Sleep(1000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });


        //                        //str = "uci set wireless.@wifi-iface[1].htmode=VHT" + smode;
        //                        str = "uci set wireless.@wifi-iface[1].htmode=VHT" + smode;
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(1000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });


        //                        /* Configure security */

        //                        /* Reset wifi */
        //                        str = "wifi down";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                        str = "wifi up";
        //                        tc.WriteLine(str);
        //                        while (true)
        //                        {

        //                            Thread.Sleep(3000);
        //                            str = tc.Read();
        //                            if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                            if (str.IndexOf("#") != -1) break;
        //                        }

        //                        str = "iwpriv wifi0 setCountryID 843";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                    }




        //                    if (smode == "80")
        //                    {
        //                        if (s[0] == "11N")
        //                        {
        //                            str = "uci set wireless.@wifi-iface[1].hwmode=11an";
        //                        }
        //                        if (s[1] == "11AC")
        //                        {
        //                            str = "uci set wireless.@wifi-iface[1].hwmode=11ac";
        //                        }

        //                        tc.WriteLine(str);
        //                        Thread.Sleep(1000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });


        //                        //str = "uci set wireless.@wifi-iface[1].htmode=VHT" + smode;
        //                        str = "uci set wireless.@wifi-iface[1].htmode=VHT" + smode;
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(1000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });


        //                        /* Configure security */

        //                        /* Reset wifi */
        //                        str = "wifi down";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                        str = "wifi up";
        //                        tc.WriteLine(str);
        //                        while (true)
        //                        {

        //                            Thread.Sleep(3000);
        //                            str = tc.Read();
        //                            if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
        //                            if (str.IndexOf("#") != -1) break;
        //                        }

        //                        str = "iwpriv wifi0 setCountryID 843";
        //                        tc.WriteLine(str);
        //                        Thread.Sleep(3000);
        //                        str = tc.Read();
        //                        Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                    }





        //                    /* Show wireless status */
        //                    str = "uci show wireless";
        //                    tc.WriteLine(str);
        //                    Thread.Sleep(3000);
        //                    str = tc.Read();
        //                    Invoke(new SetTextCallBack(SetText), new object[] { str });

        //                    /* Show speed */
        //                    str = "iwconfig";
        //                    tc.WriteLine(str);
        //                    Thread.Sleep(3000);
        //                    str = tc.Read();
        //                    Invoke(new SetTextCallBack(SetText), new object[] { str });


        //                }

        //            }

        //            // while connected   
        //                    //if (tc.IsConnected ) //&& prompt.Trim() != "exit")


        //                // send client input to server   
        //                    //  prompt = Console.ReadLine();
        //                    //  tc.WriteLine(prompt);

        //                // display server output   

        //            //    tc.WriteLine("wlanconfig ath0 list sta");
        //                    //    s = tc.Read();
        //                    //    textBox1.AppendText(s);

        //            //    tc.WriteLine("exit");



        //            //}

        //            //textBox1.AppendText("***DISCONNECTED");






        //            /* CGI command has issue, replace with telnet */
        //            // Shutdown interface 0 and 1
        //            //cgi = "http://192.168.1.1:81/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=disabled";
        //            //if (!SendCGICommand(cgi, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text, "wireless ConfigurationPass"))
        //            //{
        //            //    return false;
        //            //}

        //            //cgi = "http://192.168.1.1:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=disabled";
        //            //if (!SendCGICommand(cgi, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text, "wireless ConfigurationPass"))
        //            //{
        //            //    return false;
        //            //}

        //            //cgi = "http://" + txt_RouterSetting_DUTIPAddress.Text + ":81/mfgtst.cgi?" + "sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=1";                    
        //            //if (band == "2.4G")
        //            //{
        //            //    string iMode = mode.Substring(0, mode.Length - 1);
        //            //    if (iMode == "40")
        //            //    { /* when mode is 40M, need to switch channel to channel -2 */
        //            //        if ((channel - 2) >= 0)
        //            //            cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + (channel - 2).ToString());
        //            //        else
        //            //            cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=-1");
        //            //    }
        //            //    if (iMode == "20")
        //            //    {
        //            //        cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + channel.ToString());
        //            //    }

        //            //    cgi = cgi.Replace("sys_BandWidth=20", "sys_BandWidth=" + iMode);
        //            //    cgi = cgi.Replace("sys_wlSSID=ssid", "sys_wlSSID=" + ssid);
        //            //}

        //            ///* cgi => sys_wlInterface=0&sys_wlMode=an-mixed&sys_BandWidth=20&sys_wlSSID=HWSQA5G&sys_wlChannel=1 */
        //            //if (band == "5G")
        //            //{
        //            //    string[] s = mode.Split('_');
        //            //    if (s[0] == "11N") cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=an-mixed");
        //            //    if (s[0] == "11AC") cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=ac-mixed");
        //            //    string iMode = s[1].Substring(0, s[1].Length - 1);

        //            //    cgi = cgi.Replace("sys_wlInterface=0", "sys_wlInterface=1");
        //            //    cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=an-mixed");
        //            //    cgi = cgi.Replace("sys_BandWidth=20", "sys_BandWidth=" + iMode);
        //            //    cgi = cgi.Replace("sys_wlSSID=ssid", "sys_wlSSID=" + ssid);
        //            //    cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + channel.ToString());
        //            //}

        //            //if (!SendCGICommand(cgi, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text, "wireless ConfigurationPass"))
        //            //{
        //            //    return false;
        //            //}
        //            else return false;
        //            break;



        //        case "EA8500X8":
        //            string hostX8 = EA8500X8.host;
        //            hostX8 = hostX8.Replace("DUTIPaddr", txt_RouterSetting_DUTIPAddress.Text);
        //            string data = string.Empty;
        //            hostX8.Replace("DUTIPaddr", txt_RouterSetting_DUTIPAddress.Text);
        //            if (band == "2.4G")
        //            {
        //                string iMode = mode.Substring(0, mode.Length - 1);
                        

        //                if (iMode == "40")
        //                {
        //                    data = EA8500X8.data_2_4G_11n_40M;
        //                }
        //                if (iMode == "20")
        //                {
        //                    data = EA8500X8.data_2_4G_11n_20M;
        //                }

        //                data = data.Replace("SSIDName", ssid);
        //                data = data.Replace("channel_3", channel.ToString());
        //                if(security != "None")
        //                {
        //                    data = data.Replace(@"None""}}]}}]", security+@""",""wpaPersonalSettings"":{""passphrase"":"""+passphrase+@"""}}}]}}]");
        //                }

        //                if (!WebHttpPostEA8500X8(hostX8, data, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text))
        //                {
        //                    return false;
        //                }
        //                //WebHttpPostEA8500X8(hostX8, data, "admin", "admin");


        //            }

        //            if (band == "5G")
        //            {
        //                switch (mode)
        //                {  /* "11N_20M", "11N_40M", "11AC_20M", "11AC_40M", "11AC_80M" */
        //                    case "11N_20M":
        //                        data = EA8500X8.data_5G_11n_20M;
        //                    break ;
                            
        //                    case "11N_40M":
        //                        data = EA8500X8.data_5G_11n_40M;
        //                    break;
                            
        //                    case "11AC_20M":
        //                        data = EA8500X8.data_5G_11ac_20M;
        //                    break;
                            
        //                    case "11AC_40M":
        //                        data = EA8500X8.data_5G_11ac_40M;
        //                    break;
                            
        //                    case "11AC_80M":
        //                        data = EA8500X8.data_5G_11ac_80M;
        //                    break;

        //                    default:
        //                        return false;
        //                    break;
        //                }


        //                data = data.Replace("SSIDName", ssid);
        //                data = data.Replace("channel_3", channel.ToString());
        //                if (security != "None")
        //                {
        //                    data = data.Replace(@"None""}}]}}]", security + @""",""wpaPersonalSettings"":{""passphrase"":""" + passphrase + @"""}}}]}}]");
        //                }

        //                if (!WebHttpPostEA8500X8(hostX8, data, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text))
        //                {
        //                    return false;
        //                }
        //            }
        //            break;
                
        //        default:
        //            return false;           
        //    }

        //    return true;        
        //}
                
       

        //private string GetWiFiData()
        //{
        //    string returnData = "";            
        //    try
        //    {
        //        UdpClient receivingUdpClient = new UdpClient(9051);
        //        IPEndPoint RemoteIpEndPoint = new IPEndPoint(IPAddress.Any, 0);

        //        Byte[] receiveBytes = receivingUdpClient.Receive(ref RemoteIpEndPoint);

        //        returnData = Encoding.ASCII.GetString(receiveBytes);

        //        Debug.WriteLine("This is the message you received " + returnData.ToString());
        //        Debug.WriteLine("This message was sent from " + RemoteIpEndPoint.Address.ToString() +
        //            " on their point number" +
        //            RemoteIpEndPoint.Port.ToString());

        //        receivingUdpClient.Close();
        //        receivingUdpClient = null;
        //    }
        //    catch (Exception e)
        //    {
        //        Debug.WriteLine(e.ToString());
        //    }

        //    return returnData;
        //}

        //private void WriteCSV(string[] files, string outputfile)
        //{
        //    string value;
        //    StringBuilder csv = new StringBuilder();
        //    string newline;

        //    if (File.Exists(outputfile)) 
        //        File.Delete(outputfile);

        //    int i = 0;
        //    //for (int i = 0; i < files.Length; i++)
        //    foreach(string s in files)
        //    {
        //        //value = csv_GetSpecificValue(files[i], 15, 16);

        //         string[] Read_y_point = File.ReadAllLines(s);
        //         int y_length = Read_y_point.Length;
        //         string[] Read_y_point_value = Read_y_point[y_length - 1].Split(',');
        //         double Y_value = Convert.ToDouble(Read_y_point_value[9]);
        //         value = Y_value.ToString();


        //        newline = string.Format("{0},{1}{2}", Path.GetFileName(files[i++]), value, Environment.NewLine);
        //        csv.Append(newline);
        //    }

        //    //before your loop
        //    /*
        //    var csv = new StringBuilder();

        //    //in your loop
        //    var first = reader[0].ToString();
        //    var second = image.ToString();
        //    var newLine = string.Format("{0},{1}{2}", first, second, Environment.NewLine);
        //    csv.Append(newLine);

        //    */
        //    //after your loop
        //    File.WriteAllText(outputfile, csv.ToString());
        //}

        //private void WriteResultCSV(string files, string outputfile)
        //{
        //    //StringBuilder csv = new StringBuilder();
        //    //string newline;
        //    /*
        //    if (File.Exists(outputfile))
        //        File.Open(outputfile, FileMode.Open);
        //    */
        //    //newline = string.Format("{0},{1}{2}", files, WiFiData, Environment.NewLine);
        //    //linkrate_csv.Append(newline);

        //    //File.WriteAllText(outputfile, linkrate_csv.ToString());
        //    /*
        //    if (!File.Exists(outputfile))
        //    {
        //        FileStream filestream = File.Create(outputfile);
        //        filestream.Close();
        //    }
        //     * */
        //}

        //private string csv_GetSpecificValue(string path, int row, int col)
        //{
        //    string val = string.Empty;

        //    StreamReader file = new StreamReader(path);
        //    int count = 1;
        //    string line = string.Empty;
        //    while (!file.EndOfStream)
        //    {
        //        if (count == row)
        //        {
        //            line = file.ReadLine();
        //            val = line.Split(',')[col];
        //            file.Close();
        //            break;
        //        }
        //        file.ReadLine();
        //        count++;
        //    }
        //    return val;
        //}

        //private bool WebHttpPostEA8500X8(string head, string data, string username, string password)
        //{
        //    byte[] bs = Encoding.ASCII.GetBytes(data);

        //    //HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create("http://www.google.com/intl/zh-CN/");
        //    //HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create("http://192.168.1.1");
        //    HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(head);
        //    req.Method = "POST";
        //    req.ContentType = "application/json; charset=UTF-8";

        //    //req.ContentType = "text/xml";
        //    //req.ContentType = "application/x-www-form-urlencoded";


        //    req.ContentLength = bs.Length;
        //    //req.Headers.Add("
        //    req.Headers.Add("x-jnap-authorization", "Basic YWRtaW46YWRtaW4=\r\n");
        //    req.Headers.Add("x-jnap-action", "http://linksys.com/jnap/core/Transaction\r\n");
        //    req.Headers.Add("x-requested-with", "XMLHttpRequest\r\n");

        //    req.Headers.Add("Authorization", "Basic YWRtaW46YWRtaW4=\r\n");
        //    //req.Headers.Add("Credentials", "admin:admin");
        //    req.Headers.Add("Credentials", username + ":" + password);

        //    try
        //    {
        //        using (Stream reqStream = req.GetRequestStream())
        //        {
        //            reqStream.Write(bs, 0, bs.Length);
        //        }
        //        using (WebResponse wr = req.GetResponse())
        //        {
        //            Stream myStream = wr.GetResponseStream();
        //            StreamReader reader = new StreamReader(myStream);
        //            string strHtml = reader.ReadToEnd();
        //            //textBox1.AppendText(strHtml);

        //            //textBox1.AppendText(Environment.NewLine);
        //            //textBox1.AppendText(Environment.NewLine);


        //            if (strHtml.ToLower().IndexOf("error") >= 0)
        //            {
        //                //textBox1.AppendText("Perfect");
        //                return false;
        //            }
        //            wr.Close();
        //            return true;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //    }
        //}

        //private string ConfigureWithTelnet(string host, string username, string password)
        //{
        //    string str = string.Empty;
            



        //    return str;
        //}

        private void cbox_Model_Name_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbox_Model_Name.SelectedItem.ToString() == "EA8500X8")
            {
                //cboxThroughputTestConditionSecurity.Enabled = true;
                //txtThroughputTestConditionPassphrase.Enabled = true;
                //labThroughputTestConditionSecurity.Visible = true;
                //labThroughputTestConditionPassphrase.Visible = true;
            }
            else
            {
                //cboxThroughputTestConditionSecurity.Enabled = false;
                //txtThroughputTestConditionPassphrase.Enabled = false;
                //labThroughputTestConditionSecurity.Visible = false;
                //labThroughputTestConditionPassphrase.Visible = false;
            }
        }

        private void chkbox_RouterSetting_ConfigRouterManually_CheckedChanged(object sender, EventArgs e)
        {
            if (chkbox_RouterSetting_ConfigRouterManually.Checked)
            {
                cbox_Model_Name.Visible = false;
                label40.Visible = true;
                //txt_Temp_ModelName.Text = "";
                txt_Temp_ModelName.Visible = true;
            }
            else
            {
                cbox_Model_Name.Visible = true;
                label40.Visible = false;
                //txt_Temp_ModelName.Text = "";
                txt_Temp_ModelName.Visible = false;
            }
        }







        private void ToggleThroughputFunctionTestGUI()
        {
            ToggleThroughputFunctionTestControl(true);
            Debug.WriteLine("Toggle");
        }

        private void ToggleThroughputFunctionTestControl(bool Toggle)
        {
            /* Test Condition */
            rbtnThroughputTestCondition24G.Enabled = Toggle;
            rbtnThroughputTestCondition5G.Enabled = Toggle;
            cboxThroughputTestConditionWirelessMode.Enabled = Toggle;
            chkThroughputTestConditionWirelessSsid.Checked = Toggle;
            txtThroughputTestConditionSSIDText.Enabled = Toggle;
            chkThroughputTestConditionWirelessChannel.Checked = Toggle;
            nudThroughputTestConditionWirelessChannelStart.Enabled = Toggle;            
            nudThroughputTestConditionWirelessKeyIndex.Enabled = Toggle;
            chkThroughputTestConditionWirelessSecurity.Checked = Toggle;
            txtThroughputTestConditionPassphrase.Enabled = Toggle;
            txtThroughputTestConditionTxTst.Enabled = Toggle;
            txtThroughputTestConditionRxTst.Enabled = Toggle;
            txtThroughputTestConditionBiTst.Enabled = Toggle;

            btnThroughputTestConditionAddSetting.Enabled = Toggle;
            btnThroughputTestConditionEditSetting.Enabled = Toggle;
            btnThroughputTestConditionSaveSetting.Enabled = Toggle;
            btnThroughputTestConditionLoadSetting.Enabled = Toggle;

            /* Function Test */
            cboxThroughputFunctionTestModelName.Enabled = Toggle;
            txtThroughputFunctionTestSerialNumber.Enabled = Toggle;
            txtThroughputFunctionTestSwVersion.Enabled = Toggle;
            txtThroughputFunctionTestHwVersion.Enabled = Toggle;

            chkThroughputFunctionTestConfigRouterManually.Enabled = Toggle;
            txtThroughputFunctionTestDutModelName.Enabled = Toggle;

            txtThroughputFunctionTestDutUserName.Enabled = Toggle;
            txtThroughputFunctionTestDutPassword.Enabled = Toggle;

            txtThroughputFunctionTestDUTIPAddress.Enabled = Toggle;
            txtThroughputFunctionTest24gIPaddress.Enabled = Toggle;
            txtThroughputFunctionTest5gIPaddress.Enabled = Toggle;
            txtThroughputFunctionTestChariotFolder.Enabled = Toggle;
            nudThroughputFunctionTestPingClientTimeout.Enabled = Toggle;

            chkThroughputFunctionTestTurnTable1.Enabled = Toggle;
            nudThroughputFunctionTestTurnTable1Start.Enabled = Toggle;
            nudThroughputFunctionTestTurnTable1Stop.Enabled = Toggle;
            nudThroughputFunctionTestTurnTable1Step.Enabled = Toggle;
            txtThroughputFunctionTestTurntableUsername.Enabled = Toggle;
            txtThroughputFunctionTestTurntablePassword.Enabled = Toggle;
            txtThroughputFunctionTestTurntableClockCalibration.Enabled = Toggle;
            btnThroughputFunctionTestSaveLog.Enabled = Toggle;

            btnThroughputFunctionTestRun.Text = Toggle ? "Run" : "Stop";
            Debug.WriteLine(btnThroughputFunctionTestRun.Text);

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

            if (bThroughputFunctionTestRunning == false)
            {
                MessageBox.Show(this, "Test complete!!!", "Information", MessageBoxButtons.OK);
            }
        }
        
        /* Check all needed Condition and Attenuator is ready*/
        private bool ThroughputTestCheckCondition()
        {
            /* Check if Test Condition Data is Empty */
            if (dgvThroughputTestConditionData.RowCount <= 1)
            {
                MessageBox.Show("Configure Test Condition Data Empty, Config First!!!", "Error");
                return false;
            }

            /* Read Test Config and check if TX, RX, Bi tst file exists */
            if (!ReadTestConfig_ThroughputTest())
            {                
                SetText("Read Test Config Failed.", txtThroughputFunctionTestInformation);
                return false;
            }
           
            return true;
        }

        private bool ReadTestConfig_ThroughputTest()
        {
            string[,] rowdata = new string[dgvThroughputTestConditionData.RowCount - 1, 14];

            bTest_2_4G = false;
            bTest_5G = false;

            for (int i = 0; i < dgvThroughputTestConditionData.RowCount - 1; i++)
            {
                for (int j = 0; j < 14; j++)
                {
                    rowdata[i, j] = dgvThroughputTestConditionData.Rows[i].Cells[j].Value.ToString();
                    /* for j=0 : Set bTest_2_4G=true  if need to test 2.4G , 
                     * bTest_5G = true  if need to test 5G */
                    if (j == 0)
                    {
                        if (rowdata[i, j] == "2.4G" && !bTest_2_4G) bTest_2_4G = true;
                        if (rowdata[i, j] == "5G" && !bTest_5G) bTest_5G = true;
                    }


                    /* For j = 7~9 : Check if tst File exist */
                    if (j >= 11 && j <= 13)
                    {
                        if (rowdata[i, j] != "" && !File.Exists(rowdata[i, j]))
                        {
                            MessageBox.Show(rowdata[i, j] + " File doesn't exist. Please Check!!", "Warning");
                            return false;
                        }
                    }
                }
            }
            testConfigThroughput = rowdata;
            return true;
        }

        private bool ConfigureRouterX5(string modeType, WirelessParameter wp, int currentChannel)
        {


            return true;
        }

        private bool CreateCsvFinalReportThroughput(string filePath)
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

                str = "CyberATE Throughput + Turn Table Final Report"; //Title
                csv.AppendLine(str);

                /* Blank */
                str = "";
                csv.AppendLine(str);

                string[] _dgvHeader = new string[] { "Angle", "Band", "Mode", "Channel", "Security", "Tx", "Rx", "Bi" };
                str = string.Join(",", _dgvHeader);
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
        



        //private void ForTest()
        //{
        //    string[] a= new string[5] ;
        //    string[,] b= new string[4,5] ;

        //    int length = a.Length;
        //    int l = b.Length;
        //    int le = b.Rank;

        //    int ll = b.GetLength(0);
        //    int lf = b.GetLength(1);
        //    for (int dim = 1; dim <= le; dim++)
        //    {
        //        //string k = 
        //    }

        //    //int d = 0;
        //}
    } //End of class 
}
