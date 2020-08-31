///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : PowerOnOffFunctionTest.cs
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
using System.Diagnostics;
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
using AdamInstruments;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {           
        string m_PowerOnOffTestSubFolder;
        string[,] testConfigPowerOnOff; //test condition data
        string Address4068 = "01";
        int testTimes = 4000;
        Comport ComPortPowerOnOff = null;
        Comport2 ComPort2PowerOnOff = null;
        adam4068 adm4068 = null;
        string logFileName = string.Empty;
        int iLogTestCondition = -1;

        string startupPath = System.Windows.Forms.Application.StartupPath;
        
               
        /* Declare delegate function */
        public delegate void showPowerOnOffFunctionTestGUIDelegate();
        public delegate void AddDataGridViewDelegate(DataGridView dgv, string[] data);
        public delegate void ShowLabelTextDelegate(Label label, string str);

        private void powerOnOffTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //TestItem = "Power On Off Test";             
            sTestItem = TestItemConstants.TESTITEM_POWER_ONOFF;            
            ToggleToolStripMenuItem(false);
            tabControl_RouterStartPage.Hide();

            powerOnOffTestToolStripMenuItem.Checked = true;        

            //foreach (TabPage page in PowerOnOff_tabControl.TabPages)
            //{
            //    if (page.Text.Equals("Power On Off Test Condition", StringComparison.OrdinalIgnoreCase) || page.Text.Equals("Power On Off Function Test", StringComparison.OrdinalIgnoreCase))
            //        continue;

            //    PowerOnOff_tabControl.TabPages.Remove(page);
            //}
            PowerOnOffTabHideAllPages();

            tp_PowerOnOffTestCondition.Parent = this.PowerOnOff_tabControl; //Show tabPage
            tp_PowerOnOffFunctionTest.Parent = this.PowerOnOff_tabControl;           


            cboxPowerOnOffTestConditionAction1.SelectedIndex = 0;
            cboxPowerOnOffTestConditionAction2.SelectedIndex = 1;

            PowerOnOff_tabControl.Show();
            tsslMessage.Text = PowerOnOff_tabControl.TabPages[PowerOnOff_tabControl.SelectedIndex].Text + " Control Panel";

            string xmlFile = System.Windows.Forms.Application.StartupPath + "\\config\\" + Constants.XML_ROUTER_POWERONOFF;
            Debug.WriteLine(File.Exists(xmlFile) ? "File exist" : "File is not existing");

            if (!File.Exists(xmlFile))
            {
                WriteXmlDefaultPowerOnOff(xmlFile);
            }

            ReadXmlPowerOnOff(xmlFile);
            InitPowerOnOffTestCondition();
            InitPowerOnOffTestResult();
        }
 
        private void PowerOnOff_tabControl_Selected(object sender, TabControlEventArgs e)
        {
            if(PowerOnOff_tabControl.SelectedIndex >=0 )
                tsslMessage.Text = PowerOnOff_tabControl.TabPages[PowerOnOff_tabControl.SelectedIndex].Text + " Control Panel";
        }

        private void btnPowerOnOffFunctionTestSaveLog_Click(object sender, EventArgs e)
        {            
            SaveFileDialog saveDebugFileDialog = new SaveFileDialog();

            saveDebugFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            saveDebugFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveDebugFileDialog.FilterIndex = 1;
            saveDebugFileDialog.RestoreDirectory = true;
            saveDebugFileDialog.FileName = "Log";

            if (saveDebugFileDialog.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText(saveDebugFileDialog.FileName, txtPowerOnOffFunctionTestInformation.Text);
            }
        }

        private void btnPowerOnOffFunctionTestRun_Click(object sender, EventArgs e)
        {
            string strRead = string.Empty;

            //Thread threadRouterFT;
            //bool bRouterFTThreadRunning = false;

            /* Prevent double-click from double firing the same action */
            btnPowerOnOffFunctionTestRun.Enabled = false;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            btnPowerOnOffFunctionTestRun.Text = "Stop";

            if(btnPowerOnOffTestConditionEditSetting.Text == "Cancel")
            {
                btnPowerOnOffTestConditionEditSetting.Text = "Edit";
                hasDeleteButton = false;
                dgvPowerOnOffTestConditionData.Columns.Remove("Action");
            }

            if (bRouterFTThreadRunning == false)
            {
                //ForPoerOnOffTestOnly();
                //MessageBox.Show("For Test finished!!!");
                //btnPowerOnOffFunctionTestRun.Text = "Run";
                //btnPowerOnOffFunctionTestRun.Enabled = true;
                //System.Windows.Forms.Cursor.Current = Cursors.Default;
                //return;

                labPowerOnOffFunctionTestTimes.Text = "0";

                /* Remove all the result data */

                if (dgvPowerOnOffTestResult.RowCount > 1)
                {
                    DataTable dt = (DataTable)dgvPowerOnOffTestResult.DataSource;
                    dgvPowerOnOffTestResult.Rows.Clear();
                    dgvPowerOnOffTestResult.DataSource = dt;
                }

                if (!PowerOnOffCheckCondition())
                {
                    MessageBox.Show("Condition check failed. Please check Again!!!", "Warning");
                    btnPowerOnOffFunctionTestRun.Text = "Run";
                    btnPowerOnOffFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    return;
                }

                SetText("Check Condition OK ", txtPowerOnOffFunctionTestInformation);

                Address4068 = txtPowerOnOffFunctionTestRelayAddress.Text;
                testTimes = Convert.ToInt32(nudPowerOnOffFunctionTestTimes.Value);

                /* Create COM port */
                ComPortPowerOnOff = new Comport();

                if (ComPortPowerOnOff.isOpen() != true)
                {
                    MessageBox.Show("COM port is not ready! Please check COM port settings");
                    btnPowerOnOffFunctionTestRun.Text = "Run";
                    btnPowerOnOffFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    return;
                }

                if (chkPowerOnOffFunctionTestComportLog.Checked)
                {
                    ComPort2PowerOnOff = new Comport2();
                    if (ComPort2PowerOnOff.isOpen() != true)
                    {
                        MessageBox.Show("COM port 2 is not ready! Please check COM port 2 settings");
                        btnPowerOnOffFunctionTestRun.Text = "Run";
                        btnPowerOnOffFunctionTestRun.Enabled = true;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        return;
                    }

                    iLogTestCondition = Convert.ToInt32(nudPowerOnOffFunctionTestBindTestCondition.Value) - 1;
                }


                adm4068 = new adam4068(ComPortPowerOnOff);

                /* Create sub folder for saving the report data */
                m_PowerOnOffTestSubFolder = createPowerOnOffSubFolder(txtPowerOnOffFunctionTestName.Text);

                bRouterFTThreadRunning = true;
                /* Disable all controller */
                TogglePowerOnOffFunctionTestControl(false);
                btnPowerOnOffFunctionTestRun.Text = "Stop";


                /* Run Main Thread in according to Test Item */
                if (sTestItem == TestItemConstants.TESTITEM_POWER_ONOFF)
                {
                    threadRouterFT = new Thread(new ThreadStart(DoPowerOnOffFunctionTest));
                    threadRouterFT.Name = "";
                    threadRouterFT.Start();
                }
            }
            else
            {
                tsslMessage.Text = "Function Test Control Panel";
                bRouterFTThreadRunning = false;
            }

            /* Button release */
            Thread.Sleep(3000);
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            btnPowerOnOffFunctionTestRun.Enabled = true;
        }
         
        private void DoPowerOnOffFunctionTest()
        {
            string str = string.Empty;
            string PowerPort;
            string PowerOnTime;
            string PowerOffTime;
            string ModelName;
            string Action1;
            string Action2;
            string SleepTimer;
            string Parameter1;
            string P1_LoginID;
            string P1_LoginPW;
            string Parameter2;
            int iPowerPort = 0 ;
            int iSleepTimer = 10;    //second
            int MaxPowerOnTime = 1;  //second
            int MaxPowerOffTime = 1; //second
            int iMaxPowerOnTime = 1;
            int iMaxPowerOffTime = 1;
            int relayPort = 0;
            string relayData = "00";
            bool allPass = true;
            string ComportLogData = string.Empty;
            bool stopToDebug = false;

            string[] testFileName = new string[dgvPowerOnOffTestConditionData.RowCount-1];
            int[] TestItemStatus = new int[dgvPowerOnOffTestConditionData.RowCount-1];
            int[] TestPassCount = new int[dgvPowerOnOffTestConditionData.RowCount - 1];
            int[] TestFailCount = new int[dgvPowerOnOffTestConditionData.RowCount - 1];

            if (chkPowerOnOffFunctionTestComportLog.Checked)
            {
                ComPort2PowerOnOff.ReadLine();
                ComPort2PowerOnOff.DiscardBuffer();
            }

            /* Initial DataGridView Test Result*/
            for (int i = 0; i < dgvPowerOnOffTestConditionData.RowCount - 1; i++)
            {
                string[] data = new string[] { (i+1).ToString(), "0", "0" };
                this.Invoke(new AddDataGridViewDelegate(AddDataGridView), new object[] { dgvPowerOnOffTestResult, data });
                //dgvPowerOnOffTestResult.Rows.Add(data);
            }

            /* File name , max power on/off time , port used process */
            for (int i = 0; i < dgvPowerOnOffTestConditionData.RowCount - 1; i++)
            {
                #region
                if (bRouterFTThreadRunning == false)
                {
                    MessageBox.Show("Test Abort!!", "Warning");
                    bRouterFTThreadRunning = false;
                    this.Invoke(new showPowerOnOffFunctionTestGUIDelegate(TogglePowerOnOffFunctionTestGUI));
                    threadRouterFT.Abort();
                }

                PowerPort    = testConfigPowerOnOff[i, 0];
                ModelName    = testConfigPowerOnOff[i, 1];
                Action1      = testConfigPowerOnOff[i, 2];
                Action2      = testConfigPowerOnOff[i, 3];
                SleepTimer   = testConfigPowerOnOff[i, 4];
                Parameter1   = testConfigPowerOnOff[i, 5];
                P1_LoginID   = testConfigPowerOnOff[i, 6];
                P1_LoginPW   = testConfigPowerOnOff[i, 7];
                Parameter2   = testConfigPowerOnOff[i, 8]; 
                PowerOnTime  = testConfigPowerOnOff[i, 9];
                PowerOffTime = testConfigPowerOnOff[i, 10];                
                                
                testFileName[i] = createRouterPowerOnOffSavePath(m_PowerOnOffTestSubFolder, ModelName, Action1, Parameter1, testTimes.ToString(), 1);
                if(iLogTestCondition == i)
                {
                    if (chkPowerOnOffFunctionTestComportLog.Checked)
                    {
                        logFileName = testFileName[i].Substring(0, testFileName[i].Length - 3) + "log";
                        try
                        {
                            if (!File.Exists(logFileName))
                            {
                                FileStream logFile = File.Create(logFileName);
                                Thread.Sleep(1000);
                                logFile.Close();                               
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Log File " + logFileName + " Create Failed!!", "Warning");
                            bRouterFTThreadRunning = false;
                            this.Invoke(new showPowerOnOffFunctionTestGUIDelegate(TogglePowerOnOffFunctionTestGUI));
                            threadRouterFT.Abort();
                        }
                    }

                    ComportLogData = "Comport Log Bind Test Condition: " + (iLogTestCondition + 1).ToString() + "\r\n";
                }         
                
                
                if (!CreateCsvFilePowerOnOff(testFileName[i]))
                {
                    MessageBox.Show("CSV File " + testFileName[i] + " Create Failed!!", "Warning");
                    bRouterFTThreadRunning = false;
                    this.Invoke(new showPowerOnOffFunctionTestGUIDelegate(TogglePowerOnOffFunctionTestGUI));
                    threadRouterFT.Abort();
                }

                try
                {
                    iPowerPort = Convert.ToInt32(PowerPort);
                    relayPort += Convert.ToInt32(Math.Pow(2, iPowerPort));
                    int iPowerOnTime = Convert.ToInt32(PowerOnTime);
                    if (iPowerOnTime > MaxPowerOffTime) MaxPowerOnTime = iPowerOnTime;
                    int iPowerOffTime = Convert.ToInt32(PowerOffTime);
                    if (iPowerOffTime > MaxPowerOffTime) MaxPowerOffTime = iPowerOffTime;
                    
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Convert Power On/Off Time Error: " + ex.ToString());
                    MessageBox.Show("CSV File " + testFileName[i] + " Create Failed!!", "Warning");
                    bRouterFTThreadRunning = false;
                    this.Invoke(new showPowerOnOffFunctionTestGUIDelegate(TogglePowerOnOffFunctionTestGUI));
                    threadRouterFT.Abort();
                }

                TestPassCount[i] = 0;
                TestFailCount[i] = 0;
                #endregion
            }

            relayData = Convert.ToString(relayPort, 16);
            if (relayData.Length == 1) relayData = "0" + relayData;
            iMaxPowerOffTime = Convert.ToInt32(MaxPowerOffTime * 1e3);
            iMaxPowerOnTime = Convert.ToInt32(MaxPowerOnTime * 1e3);

            str = "Maximun Power On Time is : " + iMaxPowerOnTime.ToString();
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

            str = "Maximun Power Off Time is : " + iMaxPowerOffTime.ToString();
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

            str = "=====================Start Testing==================";
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });
            
            //======================================//                        
            //------ Turn off the relay first ------//
            //======================================//
            adm4068.TurnOffRelay(Address4068);
            Thread.Sleep(iMaxPowerOffTime);
            str = "Turn Off Relay ";
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();


            //======================================//
            //----------- Main Test Loop -----------//
            //======================================//
            for (int timesCount = 1; timesCount <= testTimes; timesCount++)
            {
                #region Main Test Loop
                /* Reset Test Status */
                for (int item = 0; item < dgvPowerOnOffTestConditionData.RowCount - 1; item++)
                {
                    TestItemStatus[item] = 0;                    
                }
                allPass = false;

                str = "====Start Test " + timesCount.ToString() + "====";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });
                ComportLogData = ComportLogData + str + "\r\n";

                this.Invoke(new ShowLabelTextDelegate(ShowLabelText), new object[] { labPowerOnOffFunctionTestTimes, timesCount.ToString() });
                //labPowerOnOffFunctionTestTimes.Text = timesCount.ToString();

                //----------------------------------//
                //*** Test Step 1: Turn on relay ***//
                //----------------------------------//
                adm4068.TurnOnRelay(Address4068, relayData);
                str = "Turn On Relay ";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });
                     
                
                stopwatch.Stop();
                stopwatch.Reset();
                stopwatch.Restart();

                str = "Wait for " + MaxPowerOnTime + " seconds";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });
                int timecount = 0;

                while (true)
                {
                    #region
                    if (bRouterFTThreadRunning == false)
                    {
                        MessageBox.Show("Test Abort", "Warning");
                        bRouterFTThreadRunning = false;
                        this.Invoke(new showPowerOnOffFunctionTestGUIDelegate(TogglePowerOnOffFunctionTestGUI));
                        threadRouterFT.Abort();
                    }
                    if(stopwatch.ElapsedMilliseconds > iMaxPowerOnTime) break;

                    if (chkPowerOnOffFunctionTestComportLog.Checked)
                    {
                        string strRead = string.Empty;
                        //strRead = ComPort2PowerOnOff.ReadLine();
                        int count = ComPort2PowerOnOff.GetBytesToRead();
                        if (count > 0) strRead = ComPort2PowerOnOff.Read(0, count);
                        //if (strRead != "")
                        //{
                        //    //Invoke(new SetTextCallBackT(SetTextC), new object[] { strRead, txtPowerOnOffFunctionTestInformation });
                        //    ComportLogData += strRead;
                        //}

                        ComportLogData += strRead;
                        //CsvAppend(logFileName, ComportLogData);
                        CsvAppend(logFileName, ComportLogData);
                        ComportLogData = "";
                    }
                    
                    if (timecount++ % 20 == 0)
                    {
                        str = ".";
                        Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtPowerOnOffFunctionTestInformation });
                    }
                    Thread.Sleep(50);
                    #endregion
                }
                
                //------------------------------------------------------//
                //*** Test Step 2: Do the "Action 1" sequencely once ***//
                //------------------------------------------------------//
                for (int pingCount = 0; pingCount < Convert.ToInt32(nudPowerOnOffFunctionTestPingTimes.Value); pingCount++)
                {
                    #region
                    if (bRouterFTThreadRunning == false)
                    {
                        MessageBox.Show("Test Abort", "Warning");
                        bRouterFTThreadRunning = false;
                        this.Invoke(new showPowerOnOffFunctionTestGUIDelegate(TogglePowerOnOffFunctionTestGUI));
                        threadRouterFT.Abort();
                    }

                    if (!chkPowerOnOffFunctionTestPingTimes.Checked) pingCount = Convert.ToInt32(nudPowerOnOffFunctionTestPingTimes.Value) + 1;
                    
                    for (int i = 0; i < dgvPowerOnOffTestConditionData.RowCount - 1; i++)
                    {
                        if (bRouterFTThreadRunning == false)
                        {
                            MessageBox.Show("Test Abort", "Warning");
                            bRouterFTThreadRunning = false;
                            this.Invoke(new showPowerOnOffFunctionTestGUIDelegate(TogglePowerOnOffFunctionTestGUI));
                            threadRouterFT.Abort();
                        }

                        /*
                        PowerPort = testConfigPowerOnOff[i, 0];
                        ModelName = testConfigPowerOnOff[i, 1];
                        Action = testConfigPowerOnOff[i, 2];
                        Parameter1 = testConfigPowerOnOff[i, 3];
                        Parameter2 = testConfigPowerOnOff[i, 4];

                        PowerOnTime = testConfigPowerOnOff[i, 5];
                        PowerOffTime = testConfigPowerOnOff[i, 6];
                        */
                        PowerPort    = testConfigPowerOnOff[i, 0];
                        ModelName    = testConfigPowerOnOff[i, 1];
                        Action1      = testConfigPowerOnOff[i, 2];
                        Action2      = testConfigPowerOnOff[i, 3];
                        SleepTimer   = testConfigPowerOnOff[i, 4];
                        Parameter1   = testConfigPowerOnOff[i, 5];
                        P1_LoginID   = testConfigPowerOnOff[i, 6];
                        P1_LoginPW   = testConfigPowerOnOff[i, 7];
                        Parameter2   = testConfigPowerOnOff[i, 8];
                        PowerOnTime  = testConfigPowerOnOff[i, 9];
                        PowerOffTime = testConfigPowerOnOff[i, 10];


                        if (chkPowerOnOffFunctionTestComportLog.Checked)
                        {
                            string strRead = string.Empty;
                            //strRead = ComPort2PowerOnOff.ReadLine();
                            int count = ComPort2PowerOnOff.GetBytesToRead();
                            if (count > 0) strRead = ComPort2PowerOnOff.Read(0, count);
                            //if (strRead != "")
                            //{
                            //    //Invoke(new SetTextCallBackT(SetTextC), new object[] { strRead, txtPowerOnOffFunctionTestInformation });
                            //    ComportLogData += strRead;
                            //}

                            ComportLogData += strRead;
                            //CsvAppend(logFileName, ComportLogData);
                            CsvAppend(logFileName, ComportLogData);
                            ComportLogData = "";
                        }

                        if (TestItemStatus[i] == 0)
                        {
                            if (Action1.ToLower() == "ping")
                            {
                                str = "Power Port " + PowerPort + "/Device Name" + ModelName + " /Ping " + Parameter1;
                                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

                                //bool bPingResult = PingClient(Parameter1, 1000);
                                 
                                int iTimeout  = Convert.ToInt32(nudPowerOnOffFunctionTestPingTimeout.Value) *1000;

                                str = "Start to Ping, Timeout: " + iTimeout.ToString();;
                                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });
                                bool bPingResult = false;

                                int timeExpire = Convert.ToInt32(stopwatch.ElapsedMilliseconds) + iTimeout;
                                
                                while (true)
                                {
                                    if(Convert.ToInt32(stopwatch.ElapsedMilliseconds) >timeExpire)
                                    {
                                        break;
                                    }

                                    if (PingClient(Parameter1, 1000))
                                    {
                                        bPingResult = true;
                                        break;
                                    }
                                }


                                //if (PingClient(Parameter1, iTimeout))
                                if (bPingResult)
                                {
                                    int timePass = Convert.ToInt32(stopwatch.ElapsedMilliseconds / 1000);
                                    str = "Time past: " + timePass.ToString();
                                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

                                    //string strPass = String.Format("{0}, {1}, {2}, {3}, {4}", "", timesCount.ToString(), "Pass", timePass.ToString(), MaxPowerOffTime.ToString());
                                    //CsvAppend(testFileName[i], strPass);
                                    TestItemStatus[i] = 1;
                                    //TestPassCount[i]++;

                                    if (iLogTestCondition == i)
                                    {
                                        string strRead = string.Empty;
                                        //strRead = ComPort2PowerOnOff.ReadLine();
                                        int count = ComPort2PowerOnOff.GetBytesToRead();
                                        if (count > 0) strRead = ComPort2PowerOnOff.Read(0, count);

                                        ComportLogData += strRead;                                        
                                        ComportLogData = ComportLogData + "\r\n";
                                        ComportLogData = ComportLogData + "Test Result is Pass\r\n";
                                        CsvAppend(logFileName, ComportLogData);  
                                        ComportLogData = "";
                                    }
                                }
                                else
                                {
                                    int timeFail = Convert.ToInt32(stopwatch.ElapsedMilliseconds / 1000);
                                    str = "Time past: " + timeFail.ToString();
                                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

                                    //string strFail = String.Format("{0}, {1}, {2}, {3}, {4}", "", timesCount.ToString(), "Fail", timeFail.ToString(), MaxPowerOffTime.ToString());
                                    //CsvAppend(testFileName[i], strFail);
                                    TestItemStatus[i] = 0;
                                    //TestFailCount[i]++;

                                    if (iLogTestCondition == i)
                                    {
                                        string strRead = string.Empty;
                                        //strRead = ComPort2PowerOnOff.ReadLine();
                                        int count = ComPort2PowerOnOff.GetBytesToRead();
                                        if (count > 0) strRead = ComPort2PowerOnOff.Read(0, count);

                                        ComportLogData += strRead;
                                        ComportLogData = ComportLogData + "\r\n";
                                        ComportLogData = ComportLogData + "Test Result is Failed\r\n";
                                        CsvAppend(logFileName, ComportLogData);
                                        ComportLogData = "";
                                    }

                                    //if (chkPowerOnOffFunctionTestStopWhenFailed.Checked)
                                    //{
                                    //    stopToDebug = true;
                                    //    break;
                                    //}
                                }
                            }
                        }
                    }
                    #endregion
                }

                for (int i = 0; i < dgvPowerOnOffTestConditionData.RowCount - 1; i++)
                {
                    #region
                    int timeNow = Convert.ToInt32(stopwatch.ElapsedMilliseconds / 1000);
                    str = "Time to record: " + timeNow.ToString(); ;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });


                    if (TestItemStatus[i] == 1)
                    {
                        TestPassCount[i]++;
                        string strPass = String.Format("{0}, {1}, {2}, {3}, {4}", "", timesCount.ToString(), "Pass", timeNow.ToString(), MaxPowerOffTime.ToString());
                        CsvAppend(testFileName[i], strPass);
                        if (iLogTestCondition == i)
                        {
                            ComportLogData = ComportLogData + "\r\n";
                            ComportLogData = ComportLogData + "Final Test Result is Pass\r\n";
                        }

                        str = "Power Port ";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });
                    }
                    else
                    {
                        TestFailCount[i]++;
                        string strFail = String.Format("{0}, {1}, {2}, {3}, {4}", "", timesCount.ToString(), "Fail", timeNow.ToString(), MaxPowerOffTime.ToString());
                        CsvAppend(testFileName[i], strFail);

                        if (iLogTestCondition == i)
                        {
                            ComportLogData = ComportLogData + "\r\n";
                            ComportLogData = ComportLogData + "Final Test Result is Failed\r\n";
                        }

                        if (chkPowerOnOffFunctionTestStopWhenFailed.Checked)
                        {
                            stopToDebug = true;
                            break;
                        }                        
                    }
                    #endregion
                }

                if (chkPowerOnOffFunctionTestComportLog.Checked)
                {
                    CsvAppend(logFileName, ComportLogData);
                    ComportLogData = "";
                }


                for (int i = 0; i < dgvPowerOnOffTestConditionData.RowCount - 1; i++)
                {
                    str = String.Format("Test Case {0}: Pass:{1}   Fail:{2}", (i + 1).ToString(), TestPassCount[i], TestFailCount[i]); 
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

                    DataGridViewRow row = dgvPowerOnOffTestResult.Rows[i];
                    row.Cells[1].Value = TestPassCount[i].ToString();
                    row.Cells[2].Value = TestFailCount[i].ToString();
                }

                if (stopToDebug)
                {
                    if (chkPowerOnOffFunctionTestComportLog.Checked)
                    {
                        string strRead = string.Empty;
                        int count = ComPort2PowerOnOff.GetBytesToRead();
                        if (count > 0) strRead = ComPort2PowerOnOff.Read(0, count);
                        ComportLogData = strRead;
                        CsvAppend(logFileName, ComportLogData);
                        ComportLogData = "";
                    }

                    MessageBox.Show("Test Fail!!! Stop For Debug !!! Test Abort!!", "Error");
                    bRouterFTThreadRunning = false;
                    this.Invoke(new showPowerOnOffFunctionTestGUIDelegate(TogglePowerOnOffFunctionTestGUI));
                    threadRouterFT.Abort();
                }
                Thread.Sleep(3000);
                
                #region old code
                //while (true)
                //{
                    //int timePassed = Convert.ToInt32(stopwatch.ElapsedMilliseconds / 1000) + 1;

                    //str = "Time Passed " + timePassed.ToString() + " seconds";
                    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

                    //if (stopwatch.ElapsedMilliseconds > iMaxPowerOnTime)
                    //{
                    //    for (int item = 0; item < dgvPowerOnOffTestConditionData.RowCount - 1; item++)
                    //    {
                    //        /* Write Fail to report */
                    //        if (TestItemStatus[item] == 0)
                    //        {
                    //            string strFail = String.Format("{0}, {1}, {2}, {3}, {4}", "", timesCount.ToString(), "Failed", timePassed.ToString(), MaxPowerOffTime.ToString());
                    //            CsvAppend(testFileName[item], strFail);
                    //            TestFailCount[item]++;
                    //        }
                    //    }

                    //    break;
                    //}

                    //if (bRouterFTThreadRunning == false)
                    //{
                    //    MessageBox.Show("Test Abort!!", "Warning");
                    //    bRouterFTThreadRunning = false;
                    //    this.Invoke(new showPowerOnOffFunctionTestGUIDelegate(TogglePowerOnOffFunctionTestGUI));
                    //    threadRouterFT.Abort();
                    //}

                    ///* Check if the test item all passed? */
                    //allPass = true;
                    //for (int count = 0; count < dgvPowerOnOffTestConditionData.RowCount - 1; count++)
                    //{                        
                    //    if (TestItemStatus[count] == 0) allPass = false;
                    //}

                    //if (allPass == true) break;

                    //for (int i = 0; i < dgvPowerOnOffTestConditionData.RowCount - 1; i++)
                    //{
                    //    PowerPort = testConfigPowerOnOff[i, 0];
                    //    ModelName = testConfigPowerOnOff[i, 1];
                    //    Action = testConfigPowerOnOff[i, 2];
                    //    Parameter1 = testConfigPowerOnOff[i, 3];
                    //    Parameter2 = testConfigPowerOnOff[i, 4];

                    //    PowerOnTime = testConfigPowerOnOff[i, 5];
                    //    PowerOffTime = testConfigPowerOnOff[i, 6];

                    //    if (TestItemStatus[i] == 0)
                    //    {
                    //        if (Action.ToLower() == "ping")
                    //        {
                    //            str = "Power Port " + PowerPort + "/Device Name" + ModelName + " /Ping " + Parameter1;
                    //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

                    //            if (PingClient(Parameter1, 1000))
                    //            {
                    //                int timePass = Convert.ToInt32(stopwatch.ElapsedMilliseconds / 1000);
                    //                string strPass = String.Format("{0}, {1}, {2}, {3}, {4}", "", timesCount.ToString(), "Pass", timePass.ToString(), MaxPowerOffTime.ToString());
                    //                CsvAppend(testFileName[i], strPass);
                    //                TestItemStatus[i] = 1;
                    //                TestPassCount[i] ++;                                    
                    //            }
                    //        }
                    //    }
                    //}
                //}
                #endregion


                //startupPath
                //------------------------------------------------------//
                //*** Test Step 3: Do the "Action 2" sequencely once ***//
                //------------------------------------------------------//
                //for (int i = 0; i < dgvPowerOnOffTestConditionData.RowCount - 1; i++)
                //{

                PowerPort    = testConfigPowerOnOff[0, 0];
                ModelName    = testConfigPowerOnOff[0, 1];
                Action1      = testConfigPowerOnOff[0, 2];
                Action2      = testConfigPowerOnOff[0, 3];
                SleepTimer   = testConfigPowerOnOff[0, 4];
                Parameter1   = testConfigPowerOnOff[0, 5];
                P1_LoginID   = testConfigPowerOnOff[0, 6];
                P1_LoginPW   = testConfigPowerOnOff[0, 7];
                Parameter2   = testConfigPowerOnOff[0, 8];
                PowerOnTime  = testConfigPowerOnOff[0, 9];
                PowerOffTime = testConfigPowerOnOff[0, 10];

                if (Action2.CompareTo("System Sleep") == 0)
                {
                    str = "";  // 空一行
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });
                    str = "Remote System Go to Sleep Mode";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

                    //----------------------//
                    //--- Write bat file ---//
                    //----------------------//
                    string batFilePath         = startupPath + @"\RemoteSystemSleep.bat";
                    string ThirdpartyToolsPath = startupPath + "\\ThirdpartyTools";
                    string PSToolsPath     = ThirdpartyToolsPath + "\\PSTools";
                    string SystemSleepFile = ThirdpartyToolsPath + "\\SystemSleep\\shutdown.bat";

                    // psexec \\172.16.20.20 -u "Aspire E15" -p 123 -c -f C:\tmpFile\shutdown.bat
                    //string RemoteCmd = "psexec \\\\" + Parameter1 + " -u \"user\" -p \"user\" -c -f " + SystemSleepFile;
                    string RemoteCmd = "psexec \\\\" + Parameter1 + " -u \"" + P1_LoginID + "\" -p \"" + P1_LoginPW + "\" -c -f " + SystemSleepFile;
                    string PSpathCmd = "Cd " + PSToolsPath;

                    System.IO.StreamWriter WriteBatFile;
                    WriteBatFile = new System.IO.StreamWriter(batFilePath);

                    WriteBatFile.WriteLine(PSpathCmd);
                    WriteBatFile.WriteLine(RemoteCmd);
                    WriteBatFile.Close();


                    //----------------------//
                    //---- Run bat file ----//
                    //----------------------//
                    Process.Start(batFilePath);
                    Thread.Sleep(1000);


                    str = "Wait for " + SleepTimer + " seconds";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });
                    iSleepTimer = Convert.ToInt32(SleepTimer) * 1000;
                    Thread.Sleep(iSleepTimer);

                }

                //}

                

                //-----------------------------------//
                //*** Test Step 4: Turn off relay ***//
                //-----------------------------------//

                str = "Turn Off Relay ";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });

                //--Turn off relay
                adm4068.TurnOffRelay(Address4068);

                if (chkPowerOnOffFunctionTestComportLog.Checked)
                {
                    string strRead = string.Empty;
                    int count = ComPort2PowerOnOff.GetBytesToRead();
                    if (count > 0) strRead = ComPort2PowerOnOff.Read(0, count);
                    ComportLogData = strRead;
                    CsvAppend(logFileName, ComportLogData);
                    ComportLogData = "";
                }
                
                Thread.Sleep(iMaxPowerOffTime);

                if (chkPowerOnOffFunctionTestComportLog.Checked)
                {
                    string strRead = string.Empty;                    
                    int count = ComPort2PowerOnOff.GetBytesToRead();
                    if (count > 0) strRead = ComPort2PowerOnOff.Read(0, count);
                    ComportLogData = strRead;                   
                    CsvAppend(logFileName, ComportLogData);
                    ComportLogData = "";                   
                }

                str = "Next Turn ";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });
                #endregion
            }

            //======================================//
            //------ Write final result to csv -----//
            //======================================//
            for (int item = 0; item < dgvPowerOnOffTestConditionData.RowCount - 1; item++)
            {
                str = String.Format("{0}, {1}, {2}, {3}, {4}", "*******", "Passed", TestPassCount[item].ToString(), "Failed", TestFailCount[item].ToString());
                CsvAppend(testFileName[item], str);

                Thread.Sleep(1000);
                str = "End of Auto Run Testing Time : " + DateTime.Now.ToString("yyyy MM dd HH:mm:ss");                
                CsvAppend(testFileName[item], str);

                str = "=====================Test Result==================";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation }); 
                       

                str = String.Format("Power Port {0} \tPassed:{1}\t\tFailed:{2}", item.ToString(), TestPassCount[item].ToString(), TestFailCount[item].ToString());
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation });                               
            }                    

            str = "=====================Test Finished==================";
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPowerOnOffFunctionTestInformation }); 

            bRouterFTThreadRunning = false;
            this.Invoke(new showPowerOnOffFunctionTestGUIDelegate(TogglePowerOnOffFunctionTestGUI));
            threadRouterFT.Abort();            
        }

        //private bool RouterPowerOnOffMainFunction(string totalPort, string Port1, string Port2, string testMode, string UpRate, string DownRate,
        //    string AttenuationStart, string AttenuationStop, string AttenuationStep)
        //{
        //    string str = string.Empty;
        //    double attStart = 0;
        //    double attStop = 1.0;
        //    double attStep = 0.5;

        //    return true;
        //}

        /* Check all needed Condition and Attenuator is ready*/
        private bool PowerOnOffCheckCondition()
        {
            /* Check if Test Condition Data is Empty */
            if (dgvPowerOnOffTestConditionData.RowCount <= 1)
            {
                MessageBox.Show("Config Test Condition Data Empty, Config First!!!");
                return false;
            }

            /* Read Test Config and check if TX, RX, Bi tst file exists */
            if (!ReadTestConfig_PowerOnOff())
            {
                SetText("Read Test Config Failed.", txtPowerOnOffFunctionTestInformation);
                return false;
            }     
            
            return true;
        }

        private bool ReadTestConfig_PowerOnOff()
        {
            string[,] rowdata = new string[dgvPowerOnOffTestConditionData.RowCount - 1, dgvPowerOnOffTestConditionData.ColumnCount];

            for (int i = 0; i < dgvPowerOnOffTestConditionData.RowCount - 1; i++)
            {
                for (int j = 0; j < dgvPowerOnOffTestConditionData.ColumnCount; j++)
                {
                    rowdata[i, j] = dgvPowerOnOffTestConditionData.Rows[i].Cells[j].Value.ToString();
                }
            }
            testConfigPowerOnOff = rowdata;
            return true;
        }

        private string createPowerOnOffSubFolder(string ModelName)
        {
            string subFolder = ((ModelName == "") ? "PowerOnOff_" : ModelName + "_") +
                    DateTime.Now.ToString("yyyy_MMdd_HHmmss");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

            return subFolder;
        }

        private string createRouterPowerOnOffSavePath(string subFolder, string modelName, string action, string parameter1, string Times, int Loop)
        {
            //string PathFile = @"\report\SUBFOLDER\MODEL_SN_SW_HW_QAMTYPE_LEVELdBmV_(POS)_DATE_(LOOP).xlsx";
            //string PathFile = @"\report\SUBFOLDER\MODEL_SN_SW_HW_P1P2_TESTMODE_Rate_UpstreamRate_DownStreamRate__Attenuation_ATTENUATINSTART_ATTENUATINSTOP_ATTENUATINSTEP_DATE_(LOOP).xlsx";
            string PathFile = @"\report\SUBFOLDER\MODELNAME_ACTION_PARAMETER1_TIMES_DATE.csv";

            PathFile = PathFile.Replace("SUBFOLDER", subFolder);            
            PathFile = PathFile.Replace("MODELNAME", modelName);
            PathFile = PathFile.Replace("ACTION", action);
            PathFile = PathFile.Replace("PARAMETER1", parameter1);
            PathFile = PathFile.Replace("TIMES", Times);           
            
            PathFile = PathFile.Replace("DATE", DateTime.Now.ToString("yyyy-MMdd-HH-mm"));
            //PathFile = PathFile.Replace("LOOP", Loop.ToString());

            PathFile = System.Windows.Forms.Application.StartupPath + PathFile;

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

            return PathFile;
        }
        
        private bool CreateCsvFilePowerOnOff(string filePath)
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
                string str = "Start Auto Run Testing Time : " + DateTime.Now.ToString("yyyy MM dd HH:mm:ss");
                csv.AppendLine(str);

                /* Blank */
                str = String.Format("{0}, {1}", "", "");
                csv.AppendLine(str);

                str = String.Format("{0}, {1}, {2}, {3}, {4}", "", "Times", "Result", "Power On", "Power Off");
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

        private void TogglePowerOnOffFunctionTestGUI()
        {
            TogglePowerOnOffFunctionTestControl(true);
            Debug.WriteLine("Toggle");
        }

        /* Disable/Enable controller */
        private void TogglePowerOnOffFunctionTestControl(bool Toggle)
        {
            nudPowerOnOffTestConditionPowerPort.Enabled = Toggle;
            nudPowerOnOffTestConditionPowerOnTime.Enabled = Toggle;
            nudPowerOnOffTestConditionPowerOffTime.Enabled = Toggle;
            txtPowerOnOffTestConditionModelName.Enabled = Toggle;
            cboxPowerOnOffTestConditionAction1.Enabled = Toggle;
            txtPowerOnOffTestConditionParameter1.Enabled = Toggle;
            txtPowerOnOffTestConditionParameter2.Enabled = Toggle;
            btnPowerOnOffTestConditionAddSetting.Enabled = Toggle;
            btnPowerOnOffTestConditionEditSetting.Enabled = Toggle;
            btnPowerOnOffTestConditionSaveSetting.Enabled = Toggle;
            btnPowerOnOffTestConditionLoadSetting.Enabled = Toggle;

            chkPowerOnOffFunctionTestComportLog.Enabled = Toggle;
            nudPowerOnOffFunctionTestBindTestCondition.Enabled = Toggle;
            chkPowerOnOffFunctionTestStopWhenFailed.Enabled = Toggle;
            chkPowerOnOffFunctionTestPingTimes.Enabled = Toggle;
            nudPowerOnOffFunctionTestPingTimes.Enabled = Toggle;
            rbtnPowerOnOffFunctionTestPingFinishToNext.Enabled = Toggle;
            rbtnPowerOnOffFunctionTestPingSequence.Enabled = Toggle;
            nudPowerOnOffFunctionTestPingTimeout.Enabled = Toggle;

            // txtPowerOnOffFunctionTestRelayAddress.Enabled = Toggle;
            txtPowerOnOffFunctionTestName.Enabled = Toggle;
            nudPowerOnOffFunctionTestTimes.Enabled = Toggle;
            btnPowerOnOffFunctionTestSaveLog.Enabled = Toggle;

            btnPowerOnOffFunctionTestRun.Text = Toggle ? "Run" : "Stop";
            Debug.WriteLine(btnPowerOnOffFunctionTestRun.Text);

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
             
        private void WriteXmlDefaultPowerOnOff(string filename)
        {
            XmlWriterSettings setting = new XmlWriterSettings();
            setting.Indent = true; //指定縮排

            XmlWriter writer = XmlWriter.Create(filename, setting);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by RvR Test program.");
            writer.WriteStartElement("CyberRouterATE");
            writer.WriteAttributeString("Item", "Power On Off Test");
                                    
            /* Write Function Test Setting */
            writer.WriteStartElement("FunctionTest");
            // Test section
            writer.WriteStartElement("TestSetting");
            writer.WriteElementString("TestName", "PowerOnOff");
            writer.WriteElementString("TestTimes", "4000");            
            writer.WriteEndElement();

            // Relay Setting
            writer.WriteStartElement("RelaySetting");
            writer.WriteElementString("Address", "01");            
            writer.WriteEndElement();

            // Debug Setting
            writer.WriteStartElement("DebugSetting");
            writer.WriteElementString("EnableLog", "0");
            writer.WriteElementString("BindCondition", "1");
            writer.WriteElementString("EnableStop", "0");
            writer.WriteEndElement();


            /* End of Write FunctionTest*/
            writer.WriteEndElement();

            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }

        private void WriteXmlPowerOnOff(string filename)
        {
            XmlWriterSettings setting = new XmlWriterSettings();
            setting.Indent = true; //指定縮排

            XmlWriter writer = XmlWriter.Create(filename, setting);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by RvR Test program.");
            writer.WriteStartElement("CyberRouterATE");
            writer.WriteAttributeString("Item", "Power On Off Test");

            /* Write Function Test Setting */
            writer.WriteStartElement("FunctionTest");
            // Test section
            writer.WriteStartElement("TestSetting");
            writer.WriteElementString("TestName", txtPowerOnOffFunctionTestName.Text);
            writer.WriteElementString("TestTimes", nudPowerOnOffFunctionTestTimes.Value.ToString());
            writer.WriteEndElement();

            // Relay Setting
            writer.WriteStartElement("RelaySetting");
            writer.WriteElementString("Address", txtPowerOnOffFunctionTestRelayAddress.Text);
            writer.WriteEndElement();

            // Debug Setting
            writer.WriteStartElement("DebugSetting");
            writer.WriteElementString("EnableLog", chkPowerOnOffFunctionTestComportLog.Checked? "1" : "0");
            writer.WriteElementString("BindCondition", nudPowerOnOffFunctionTestBindTestCondition.Value.ToString());
            writer.WriteElementString("EnableStop", chkPowerOnOffFunctionTestStopWhenFailed.Checked? "1" : "0");
            writer.WriteEndElement();

            /* End of Write FunctionTest*/
            writer.WriteEndElement();

            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }

        private bool ReadXmlPowerOnOff(string filename)
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
            if (strID.CompareTo("Power On Off Test") != 0)
            {
                MessageBox.Show("This XML file is incorrect", "Error");
                return false;
            }            

            /* Read Function Test Setting */
            /* Read TestSetting */
            XmlNode nodeModel = doc.SelectSingleNode("/CyberRouterATE/FunctionTest/TestSetting");
            try
            {
                string TestName = nodeModel.SelectSingleNode("TestName").InnerText;
                string TestTimes = nodeModel.SelectSingleNode("TestTimes").InnerText;          

                
                Debug.WriteLine("Test Name: " + TestName);
                Debug.WriteLine("Test Times: " + TestTimes);

                txtPowerOnOffFunctionTestName.Text = TestName;
                nudPowerOnOffFunctionTestTimes.Value = Convert.ToDecimal(TestTimes);               
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/FunctionTest/TestSetting " + ex);
            }

            /* Read Relay Setting */
            XmlNode nodeRouter = doc.SelectSingleNode("/CyberRouterATE/FunctionTest/RelaySetting");
            try
            {
                string Address = nodeRouter.SelectSingleNode("Address").InnerText;                

                Debug.WriteLine("Address: " + Address);

                txtPowerOnOffFunctionTestRelayAddress.Text = Address;                
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/FunctionTest/RelaySetting " + ex);
            }

            /* Read Debug Setting */
            XmlNode nodeDebugSetting = doc.SelectSingleNode("/CyberRouterATE/FunctionTest/DebugSetting");
            try
            {
                string EnableLog = nodeDebugSetting.SelectSingleNode("EnableLog").InnerText;
                string BindCondition = nodeDebugSetting.SelectSingleNode("BindCondition").InnerText;
                string EnableStop = nodeDebugSetting.SelectSingleNode("EnableStop").InnerText;

                Debug.WriteLine("EnableLog: " + EnableLog);
                Debug.WriteLine("BindCondition: " + BindCondition);
                Debug.WriteLine("EnableStop: " + EnableStop);

               
                chkPowerOnOffFunctionTestComportLog.Checked = (EnableLog == "1"? true:false);
                nudPowerOnOffFunctionTestBindTestCondition.Value = Convert.ToInt32(BindCondition);
                chkPowerOnOffFunctionTestStopWhenFailed.Checked = (EnableStop == "1"? true:false);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CyberRouterATE/FunctionTest/DebugSetting " + ex);
            }


            return true;
        }

        private void PowerOnOffTabHideAllPages()
        {
            tp_PowerOnOffTestCondition.Parent = null; //Hide tabPage
            tp_PowerOnOffFunctionTest.Parent = null;            
        }
        
        private void InitPowerOnOffTestResult()
        {
            dgvPowerOnOffTestResult.ColumnCount = 3;
            dgvPowerOnOffTestResult.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvPowerOnOffTestResult.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvPowerOnOffTestResult.Name = "Power OnOff Test Result";
            //dgvPowerOnOffTestConfitionData.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgvPowerOnOffTestResult.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvPowerOnOffTestResult.Columns[0].Name = "Test Case";
            dgvPowerOnOffTestResult.Columns[1].Name = "Pass";
            dgvPowerOnOffTestResult.Columns[2].Name = "Fail";
            

            dgvPowerOnOffTestResult.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPowerOnOffTestResult.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvPowerOnOffTestResult.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False; //標題列換行, true -換行, false-不換行
            dgvPowerOnOffTestResult.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//標題列置中

            dgvPowerOnOffTestResult.Columns[0].Width = 100;
            dgvPowerOnOffTestResult.Columns[1].Width = 100;
            dgvPowerOnOffTestResult.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dgvPowerOnOffTestResult.Columns[3].Width = 120;
            //dgvPowerOnOffTestResult.Columns[4].Width = 120;
            //dgvPowerOnOffTestResult.Columns[5].Width = 100;
            //dgvPowerOnOffTestResult.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dgvPowerOnOffTestResult.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            //dgvPowerOnOffTestConfitionData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; //Use the column width setting
            dgvPowerOnOffTestResult.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            /*
             http://blog.csdn.net/alisa525/article/details/7556771
             // 設定包括Header和所有儲存格的列寬自動調整
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;  //設置列標題不換行
             // 設定不包括Header所有儲存格的行高自動調整
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;  //AllCells;設定包括Header和所有儲存格的行高自動調整            
             */





            //string[] data = new string[] { "2", "15", "30" };
            //dgvPowerOnOffTestResult.Rows.Add(data);

            //DataGridViewRow row = dgvPowerOnOffTestResult.Rows[0];

            //row.Cells[1].Value = "20";



        }

        public void AddDataGridView(DataGridView dgv, string[] data)
        {
            dgv.Rows.Add(data);            
        }

        public void ShowLabelText(Label label, string str)
        {
            label.Text = str;
        }
            



       
        private void ForPoerOnOffTestOnly()
        {
            string txtInfo = string.Empty;

            SetText("For Test", txtPowerOnOffFunctionTestInformation);
            //Comport com = new Comport();
            //if (!com.isOpen())
            //{
            //    MessageBox.Show("The Comport Opens Failed", "Warning");
            //    return;
            //}

            //adam4068 a4068 = new adam4068(com);
            //txtInfo = a4068.GetRelayStatus("01");

            //SetText(txtInfo, txtPowerOnOffFunctionTestInformation);

            //txtInfo = a4068.GetFirmwareVersion("01");

            //SetText(txtInfo, txtPowerOnOffFunctionTestInformation);

            string value = SByte.MinValue.ToString("X");
            
            int test_mode = 0;

            test_mode += 1;
            test_mode += 8;
            test_mode += 4;
            test_mode += 2;
            //test_mode += 64;
            
            string x = test_mode.ToString();

            //byte number = Convert.ToByte(x, 16);
            //SetText(number.ToString(), txtPowerOnOffFunctionTestInformation);
            //SetText(number.ToString(), txtPowerOnOffFunctionTestInformation);
            
            //txtPowerOnOffFunctionTestInformation.Text = number.ToString();

            //byte number = Convert.ToByte(test_mode, 16);
            //Console.WriteLine("0x{0} converts to {1}.", test_mode, number);
            string y = Convert.ToString(test_mode, 16);

            SetText(y, txtPowerOnOffFunctionTestInformation);


        }
    } //End of class 
}
