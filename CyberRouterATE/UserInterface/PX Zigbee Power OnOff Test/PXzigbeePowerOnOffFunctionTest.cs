///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Sally Lee.
///  File           : PXzigbeePowerOnOffFunctionTest.cs
///  Update         : 2020-08-11
///  Modified       : 2020-08-11 Initial version  
///                   
///---------------------------------------------------------------------------------------

#define DEBUG_MODE
#undef DEBUG_MODE


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
        public delegate void AddDataGridViewDelegate_PXzigbeePowerOnOff(int iROW, int iCELL, string dataValue);
        private delegate void PXzigbeePowerOnOffSetTextCallBack(string text, TextBox textbox);
        public delegate void showPXzigbeePowerOnOffDelegate();


        System.IO.StreamWriter PXzigbeePowerOnOffWriteLog;
       

        string[,] testConfigPXzigbeePowerOnOff;  // Save Test Condition Data
        PXzigbeePowerOnOff_TestConfig[] structPXzigbeePowerOnOffConfig;
        PXzigbeePowerOnOff_TestSetting structPXzigbeePowerOnOffTestSetting = new PXzigbeePowerOnOff_TestSetting();
        Thread threadPXzigbeePowerOnOffFT;
        Thread threadPXzigbeePowerOnOffFTstopEvent;
        Thread threadPXzigbeePowerOnOffMainFunction;
        Thread threadPXzigbeePowerOnOffGetStatus;
        bool bPXzigbeePowerOnOffFTThreadRunning = false;
        bool bPXzigbeePowerOnOffTestComplete = true;

        int iTestFinishCount = 0;    // Test item Finish Counter
        int iTestItemsCount  = 0;    // "DataGridView" Row Data Counter
        Dictionary<string, PXzigbeePowerOnOff_GetLightStatusLog> DictGetLightStatusLog = new Dictionary<string, PXzigbeePowerOnOff_GetLightStatusLog>();


        struct PXzigbeePowerOnOff_GetLightStatusLog
        {
            public List<string> LogList;
            public string[] LogArray;
            public int LogArrayLogIndex;
            public int LogArrayLastSaveIndex;
            public bool writeLogFinished;
        }

        struct PXzigbeePowerOnOff_TestConfig
        {
            public string ThreadName;
            public string ModeName;
            public string IP;
            public string NodeID;
            public string Action1;
            public string Action2;
            public string Action3;
            public string Action4;
            public string SleepTimer1;
            public string SleepTimer2;
            public string SleepTimer3;
            public string SleepTimer4;
            public string MQTT_Path;
        }

        struct PXzigbeePowerOnOff_TestSetting
        {
            public string TestName;
            public string TestTimes;
            public bool StopWhenTestFailed;
        }



        //**********************************************************************************//
        //------------------- PX Zigbee Power OnOff Function Test Event --------------------//
        //**********************************************************************************//
#region PX Zigbee Power OnOff Function Test Event
        private void pxZigbeePowerOnOffTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //----- Hide All TabPage -----//
            ToggleToolStripMenuItem(false);
            this.SallyTestPage.Parent = null;     // hide SallyTestPage TabPage
            tp_PXzigbeePowerOnOffTestCondition.Parent = null; //Hide tabPage
            tp_PXzigbeePowerOnOffFunctionTest.Parent  = null;            
            tabControl_RouterStartPage.Hide();

            //--- Show TabPage(needed) ---//
            PXzigbeePowerOnOff_tabControl.Show();
            this.tp_PXzigbeePowerOnOffTestCondition.Parent = this.PXzigbeePowerOnOff_tabControl;  //show the Test Condition TabPage
            this.tp_PXzigbeePowerOnOffFunctionTest.Parent  = this.PXzigbeePowerOnOff_tabControl;  //show the Function Test TabPage

            pxZigbeePowerOnOffTestToolStripMenuItem.Checked = true;

            // Default Selection
            cboxPXzigbeePowerOnOffTestConditionAction1.SelectedIndex = 2; // Power Off
            cboxPXzigbeePowerOnOffTestConditionAction2.SelectedIndex = 3; // Get Light Status
            cboxPXzigbeePowerOnOffTestConditionAction3.SelectedIndex = 1; // Power On
            cboxPXzigbeePowerOnOffTestConditionAction4.SelectedIndex = 3; // Get Light Status

            InitPXzigbeePowerOnOffTestCondition();

            //ElementVisibleSetting();
            tsslMessage.Text = PXzigbeePowerOnOff_tabControl.TabPages[PXzigbeePowerOnOff_tabControl.SelectedIndex].Text + " Control Panel";

        }

        private void PXzigbeePowerOnOff_tabControl_Selected(object sender, TabControlEventArgs e)
        {
            if (PXzigbeePowerOnOff_tabControl.SelectedIndex >= 0)
                tsslMessage.Text = PXzigbeePowerOnOff_tabControl.TabPages[PXzigbeePowerOnOff_tabControl.SelectedIndex].Text + " Control Panel";
        }

        private void btnPXzigbeePowerOnOffFunctionTestRun_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Run Btn Checked!!");
            TogglePXzigbeePowerOnOffFunctionTestSystemCursorWait();

            if (bPXzigbeePowerOnOffFTThreadRunning == false && bPXzigbeePowerOnOffTestComplete == true)
            {
                #region
                txtPXzigbeePowerOnOffFunctionTestInformation.Clear();

                //-------------- Initial & Save Test Parameter --------------//
                PXzigbeePowerOnOffInitialParameter();
                PXzigbeePowerOnOffSaveTestSetting();

                /* Remove all the result data */
                //if (dgvPowerOnOffTestResult.RowCount > 1)
                //{
                //    DataTable dt = (DataTable)dgvPowerOnOffTestResult.DataSource;
                //    dgvPowerOnOffTestResult.Rows.Clear();
                //    dgvPowerOnOffTestResult.DataSource = dt;
                //}

                if (!PXzigbeePowerOnOffCheckCondition())
                {
                    MessageBox.Show("Condition check failed. Please check Again!!!", "Warning");
                    TogglePXzigbeePowerOnOffFunctionTestSystemCursorDefault();
                    return;
                }
                SetText("Check Test Condition OK.", txtPXzigbeePowerOnOffFunctionTestInformation);
                SetText("", txtPXzigbeePowerOnOffFunctionTestInformation);



                //-------------------------------------------------//
                //----------------- Start Test --------------------//
                //-------------------------------------------------//
                bPXzigbeePowerOnOffFTThreadRunning = true;
                bPXzigbeePowerOnOffTestComplete = false;
                TogglePXzigbeePowerOnOffFunctionTestController(false);  /* Disable all controller */

                //--------------- Start Test Thread ---------------//
                SetText("----------------------------- Start PX Zigbee Power On/Off ATE Test -----------------------------", txtPXzigbeePowerOnOffFunctionTestInformation);
                threadPXzigbeePowerOnOffFT = new Thread(new ThreadStart(DoPXzigbeePowerOnOffFunctionTest));
                threadPXzigbeePowerOnOffFT.Name = "";
                threadPXzigbeePowerOnOffFT.Start();

                //*** Use another thread to catch the stop event of test thread ***//
                threadPXzigbeePowerOnOffFTstopEvent = new Thread(new ThreadStart(threadPXzigbeePowerOnOffFT_StopEvent));
                threadPXzigbeePowerOnOffFTstopEvent.Name = "";
                threadPXzigbeePowerOnOffFTstopEvent.Start();
                #endregion
            }
            else if (bPXzigbeePowerOnOffFTThreadRunning == true && bPXzigbeePowerOnOffTestComplete == false)
            {
                tsslMessage.Text = "Function Test Control Panel";
                bPXzigbeePowerOnOffFTThreadRunning = false;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "The Test will stop later. Please wait!", "ATE Information", MessageBoxButtons.OK);
            }
        } 
#endregion


        //**********************************************************************************//
        //------------------- PX Zigbee Power OnOff Function Test Module -------------------//
        //**********************************************************************************//
#region PX Zigbee Power OnOff Test Module

        //=================================================================//
        //------------------------- Initial Function ----------------------//
        //=================================================================//
        #region Initial Function
        private bool PXzigbeePowerOnOffCheckCondition()
        {
            if (!SaveTestConfig_PXzigbeePowerOnOff())
            {
                SetText("Read Test Config Failed.", txtPowerOnOffFunctionTestInformation);
                return false;
            }
            return true;
        }

        private void PXzigbeePowerOnOffInitialParameter()
        {
            iTestFinishCount = 0;
            iTestItemsCount  = 0;
            testConfigPXzigbeePowerOnOff   = null;
            structPXzigbeePowerOnOffConfig = null;
            DictGetLightStatusLog.Clear();
        }
        #endregion //-- Initial Function


        //=================================================================//
        //----------------------- Main Test Structure ---------------------//
        //=================================================================//
        #region Main Test Structure

        private void threadPXzigbeePowerOnOffFT_StopEvent()
        {
            while (bPXzigbeePowerOnOffFTThreadRunning == true)
            {
                Thread.Sleep(500);
            }
            PXzigbeePowerOnOffFunctionTestThreadFinished();
        }

        private void PXzigbeePowerOnOffFunctionTestThreadFinished()
        {
            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { "------------------------------ PX Zigbee Power On/Off ATE Test End ------------------------------", txtPXzigbeePowerOnOffFunctionTestInformation });
            this.Invoke(new showPXzigbeePowerOnOffDelegate(TogglePXzigbeePowerOnOffFunctionTestSettingTrue));
        }

        private void TestFinishedActionPXzigbeePowerOnOff()
        {
            //PXzigbeePowerOnOffWriteLog.Close();
            //CloseExcelReportFile(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade);
            //ClodeWebDriver();
            //Thread.Sleep(3000);
        }

        private void DoPXzigbeePowerOnOffFunctionTest()
        {
            //MessageBox.Show("DoPXzigbeePowerOnOffFunctionTest()");

            //this.Invoke(new showPXzigbeePowerOnOffDelegate(TogglePXzigbeePowerOnOffFunctionTestSystemCursorWait));
            //Thread.Sleep(2000);


            //this.Invoke(new showPXzigbeePowerOnOffDelegate(TogglePXzigbeePowerOnOffTestTimes));
            this.Invoke(new showPXzigbeePowerOnOffDelegate(TogglePXzigbeePowerOnOffFunctionTestSystemCursorDefault));
            //Thread.Sleep(10000);

            //string str = "iTestItemsCount: " + iTestItemsCount;
            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtPXzigbeePowerOnOffFunctionTestInformation });
            string NodeID = "";
            string MQTT_Path = "";
            string MQTT_pub = "";
            string errMsg = "";
            for (int iIndex = 0; iIndex < iTestItemsCount; iIndex++)
            {
                errMsg = "";
                NodeID = structPXzigbeePowerOnOffConfig[iIndex].NodeID;
                MQTT_Path = structPXzigbeePowerOnOffConfig[iIndex].MQTT_Path;
                MQTT_pub = String.Format("\"{0}\\mosquitto_pub.exe\"", MQTT_Path);

                try
                {
                    Process.Start(MQTT_pub);
                }
                catch (Exception ex)
                {
                    if (ex.ToString().Contains("0x80004005"))
                    {
                        errMsg = String.Format("Error File Path (node id {0}):\n{1}", NodeID, MQTT_Path);
                        MessageBoxTopMost("ATE Information", errMsg);
                    }
                    else
                        MessageBoxTopMost("Error", ex.ToString());
                    continue;
                }
                
                threadPXzigbeePowerOnOffMainFunction = new Thread(new ThreadStart(PXzigbeePowerOnOffMainFunction));
                threadPXzigbeePowerOnOffMainFunction.Name = "Node-" + (iIndex+1).ToString();
                structPXzigbeePowerOnOffConfig[iIndex].ThreadName = threadPXzigbeePowerOnOffMainFunction.Name;
                
                threadPXzigbeePowerOnOffMainFunction.Start();
                Thread.Sleep(1000);
            }

            while (iTestFinishCount != iTestItemsCount)
            {
                Thread.Sleep(100);
            }
            bPXzigbeePowerOnOffFTThreadRunning = false;
        }

        private void PXzigbeePowerOnOffMainFunction()
        {   //Thread.CurrentThread.Abort
            //Thread.CurrentThread.Join

            string CurThreadName = Thread.CurrentThread.Name;
            string Info = "";
            string strActionName = "";
            //string tempOutput = "";
            //string output     = "";
            string previousResult = "";
            string currentResult  = "";
            int iCurLoop = 0;
            int iCurIndex = 0;
            int TestTimes = Convert.ToInt32(structPXzigbeePowerOnOffTestSetting.TestTimes);
            int strSleepTimer = 1000;
            PXzigbeePowerOnOff_GetLightStatusLog structNode_tmp = new PXzigbeePowerOnOff_GetLightStatusLog();
            string[] Last50LogArray = new string[50];

            #region
            /*
            //--------------< Get Light Status Process >--------------//
            Process StatusProcess = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = false;
            startInfo.FileName = @"C:\Windows\System32" + "\\cmd.exe";
            startInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;

            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            StatusProcess.StartInfo = startInfo;
            StatusProcess.Start();
            //--------------------------------------------------------//
            */
            #endregion

            //-- Find structure index for current thread (current test item) --//
            for (int iIndex = 0; iIndex < iTestItemsCount; iIndex++)
            {
                if (structPXzigbeePowerOnOffConfig[iIndex].ThreadName.CompareTo(CurThreadName) == 0)
                {
                    iCurIndex = iIndex;
                    break;
                }
            }

            string ip = structPXzigbeePowerOnOffConfig[iCurIndex].IP;
            string id = structPXzigbeePowerOnOffConfig[iCurIndex].NodeID;
            string MQTT_Path = structPXzigbeePowerOnOffConfig[iCurIndex].MQTT_Path;
            string TAG = String.Format(" {0} | {1} |", CurThreadName, id);

            DateTime myDate = DateTime.Now;
            string myDateString = myDate.ToString("yyyyMMdd-HHmmss");
            string PXzigbeePowerOnoffFileName = String.Format(@"\PXzigbeePowerOnoff_GetStatusLog_{0}_{1}.txt", id, myDateString);
            string PXzigbeePowerOnoffFilePath = System.Windows.Forms.Application.StartupPath + PXzigbeePowerOnoffFileName;
            PXzigbeePowerOnOffWriteLog = new System.IO.StreamWriter(PXzigbeePowerOnoffFilePath);
            PXzigbeePowerOnOffWriteLog.AutoFlush = true;

                                              // mosquitto_pub -h 192.168.18.1 -t "gw/commands" -q 2 -m "{\"commands\":[{\"commandcli\":\"zcl global read 0x0006 0x0000\"},{\"commandcli\":\"send 0x1F57 1 0xff\"}]}"
            //string StatusCmd   = String.Format("mosquitto_pub -h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl global read 0x0006 0x0000\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xff\\\"}}]}}\"", ip, id);

                                              // mosquitto_pub -h 192.168.18.1 -t "gw/commands" -q 2 -m "{\"commands\":[{\"commandcli\":\"zcl on-off on\"},{\"commandcli\":\"send 0x1F57 1 0xFF\"}]}"
            //string PowerOnCmd  = String.Format("mosquitto_pub -h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl on-off on\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xFF\\\"}}]}}\"", ip, id);
            string PowerOnCmd = String.Format("-h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl on-off on\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xFF\\\"}}]}}\"", ip, id);

                                              // mosquitto_pub -h 192.168.18.1 -t "gw/commands" -q 2 -m "{\"commands\":[{\"commandcli\":\"zcl on-off off\"},{\"commandcli\":\"send 0x1F57 1 0xFF\"}]}"
            //string PowerOffCmd = String.Format("mosquitto_pub -h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl on-off off\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xFF\\\"}}]}}\"", ip, id);
            string PowerOffCmd = String.Format("-h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl on-off off\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xFF\\\"}}]}}\"", ip, id);

            //string MQTT_PathCmd     = String.Format("cd {0}", MQTT_Path);                                    // cd C:\Program Files\Project X\mos158
            //string MQTT_responseCmd = String.Format("mosquitto_sub -h {0} -t \"gw/zclresponse\" -q 2", ip);  // mosquitto_sub -h 192.168.18.1 -t \"gw/zclresponse\" -q 2



            //-------------------< Power On Process Setting>-----------------//
            Process PowerOnProcess = new Process();
            ProcessStartInfo PowerOnstartInfo = new ProcessStartInfo();
            PowerOnstartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            PowerOnstartInfo.CreateNoWindow = true;
            PowerOnstartInfo.FileName = MQTT_Path + "\\mosquitto_pub.exe";
            PowerOnstartInfo.Arguments = PowerOnCmd;
            PowerOnstartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
            PowerOnstartInfo.UseShellExecute = false;
            PowerOnstartInfo.RedirectStandardOutput = true;
            PowerOnstartInfo.RedirectStandardError = true;
            PowerOnstartInfo.RedirectStandardInput = true;
            PowerOnProcess.StartInfo = PowerOnstartInfo;
            //---------------------------------------------------------------//

            //------------------< Power Off Process Setting>-----------------//
            Process PowerOffProcess = new Process();
            ProcessStartInfo PowerOffstartInfo = new ProcessStartInfo();
            PowerOffstartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            PowerOffstartInfo.CreateNoWindow = true;
            PowerOffstartInfo.FileName = MQTT_Path + "\\mosquitto_pub.exe";
            PowerOffstartInfo.Arguments = PowerOffCmd;
            PowerOffstartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
            PowerOffstartInfo.UseShellExecute = false;
            PowerOffstartInfo.RedirectStandardOutput = true;
            PowerOffstartInfo.RedirectStandardError = true;
            PowerOffstartInfo.RedirectStandardInput = true;
            PowerOffProcess.StartInfo = PowerOffstartInfo;
            //---------------------------------------------------------------//


            //PowerOnProcess.StandardInput.WriteLine("C:");
            //PowerOnProcess.StandardInput.WriteLine(MQTT_PathCmd);

            //PowerOffProcess.StandardInput.WriteLine("C:");
            //PowerOffProcess.StandardInput.WriteLine(MQTT_PathCmd);

            List<string> ListAction = new List<string>
            {
                structPXzigbeePowerOnOffConfig[iCurIndex].Action1,
                structPXzigbeePowerOnOffConfig[iCurIndex].Action2,
                structPXzigbeePowerOnOffConfig[iCurIndex].Action3,
                structPXzigbeePowerOnOffConfig[iCurIndex].Action4
            };
            List<string> ListSleepTimer = new List<string>
            {
                structPXzigbeePowerOnOffConfig[iCurIndex].SleepTimer1,
                structPXzigbeePowerOnOffConfig[iCurIndex].SleepTimer2,
                structPXzigbeePowerOnOffConfig[iCurIndex].SleepTimer3,
                structPXzigbeePowerOnOffConfig[iCurIndex].SleepTimer4
            };


            if (ListAction.Contains("Get Light Status") == true)
            {
                #region
                //Process StatusProcess = new Process();
                //ProcessStartInfo startInfo = new ProcessStartInfo();
                //startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                //startInfo.CreateNoWindow = true;
                //startInfo.FileName = @"C:\Windows\System32" + "\\cmd.exe";
                //startInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
                //startInfo.UseShellExecute = false;
                //startInfo.RedirectStandardOutput = true;
                //startInfo.RedirectStandardError = true;
                //startInfo.RedirectStandardInput = true;
                //StatusProcess.StartInfo = startInfo;
                //StatusProcess.Start();
                

                /*
                StatusProcess.StandardInput.WriteLine("C:");
                StatusProcess.StandardInput.WriteLine(MQTT_PathCmd);
                StatusProcess.StandardInput.WriteLine(StatusCmd);
                StatusProcess.StandardInput.WriteLine(MQTT_responseCmd);

                while ((tempOutput = StatusProcess.StandardOutput.ReadLine()) != null || (tempOutput = StatusProcess.StandardError.ReadLine()) != null)
                {
                    //output += tempOutput + "\n";
                    output = tempOutput + "\n";
                    PXzigbeePowerOnOffWriteLog.WriteLine(output);
                    if (output.Contains("gw/zclresponse") == true)
                        break;
                }
                */

                //PXzigbeePowerOnOffWriteLog.Close();
                //StatusProcess.WaitForExit();
                //StatusProcess.Close();
                #endregion

                threadPXzigbeePowerOnOffGetStatus = new Thread(new ParameterizedThreadStart(DoPXzigbeePowerOnOffGetLightStatus));
                threadPXzigbeePowerOnOffGetStatus.Name = "GetStatusThread";
                threadPXzigbeePowerOnOffGetStatus.Start(CurThreadName);

            }


            
            //**********************************************************************//
            //**************************  Main Test Loop  **************************//
            //**********************************************************************//
            int LogArrayLastSaveIndex = -1;
            int LogArrayLastReadIndex = -1;
            int LogArrayPreReadIndex  = -1;
            int LogArrayLogIndex = -1;
            for (int t = 0; t < TestTimes; t++)
            {
                #region Main Test Loop

                bool bResult  = false;
                bool bTestFail = false;
                Stopwatch PowerOnOffTestTimer = new Stopwatch();
                iCurLoop = t + 1;
                Info = String.Format("{0} ------------------< Loop: {1} >------------------", TAG, iCurLoop.ToString());
                Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                PXzigbeePowerOnOffWriteLog.WriteLine(Info);

                //----------------------------------------------------------------//
                //   Execution step:                                              //
                //             Action 1 --> Action 2 --> Action 3--> Action 4     //
                //----------------------------------------------------------------//
                for (int a = 0; a < 4; a++)
                {
                    #region Execution Step Loop

                    strActionName = ListAction[a];
                    strSleepTimer = Convert.ToInt32(Convert.ToSingle(ListSleepTimer[a]) * 1000);
                    bTestFail = false;
                    switch (strActionName)
                    {
                        case "Power On":
                            #region
                            Info = String.Format("{0} Power On.", TAG);
                            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                            PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                            Info = String.Format("{0} Command: mosquitto_pub {1}", TAG, PowerOnCmd);
                            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                            PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                            PowerOnProcess.Start();
                            PowerOnProcess.WaitForExit();
                            PowerOnProcess.Close();
                            //PowerOnProcess.StandardInput.WriteLine(PowerOnCmd);
                            break;
                            #endregion

                        case "Power Off":
                            #region
                            Info = String.Format("{0} Power Off.", TAG);
                            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                            PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                            Info = String.Format("{0} Command: mosquitto_pub {1}", TAG, PowerOffCmd);
                            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                            PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                            PowerOffProcess.Start();
                            PowerOffProcess.WaitForExit();
                            PowerOffProcess.Close();
                            //PowerOffProcess.StandardInput.WriteLine(PowerOffCmd);
                            break;
                            #endregion

                        case "Get Light Status":
                            #region
                            Info = String.Format("{0} Get Light Status...", TAG);
                            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                            PXzigbeePowerOnOffWriteLog.WriteLine(Info);

                            if (iCurLoop == 1)
                                Thread.Sleep(300);  // 稍等一下再分析訊息

                            /* // List
                            int StatusListLastIndex = DictGetLightStatusLog[CurThreadName].LogList.Count -1;
                            string strStatus = DictGetLightStatusLog[CurThreadName].LogList[StatusListLastIndex];
                            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { strStatus, txtPXzigbeePowerOnOffFunctionTestInformation });
                            PXzigbeePowerOnOffWriteLog.WriteLine(strStatus);
                            */

                            currentResult = "";
                            bool bGetData = false;
                            bool bDataIsLastOne = true;
                            string strStatus = "";
                            int newLastSaveIndex = DictGetLightStatusLog[CurThreadName].LogArrayLastSaveIndex;
                            int newLastLogIndex  = DictGetLightStatusLog[CurThreadName].LogArrayLogIndex;
                            strStatus = String.Format("{0} Status {1}: {2}", TAG, newLastLogIndex, DictGetLightStatusLog[CurThreadName].LogArray[newLastSaveIndex]);
                            Last50LogArray[newLastSaveIndex] = strStatus;


                            #region
                            //if (LogArrayLastSaveIndex != LogArrayLastReadIndex)
                            //{
                            //}
                            //else if (LogArrayLastSaveIndex == LogArrayLastReadIndex)
                            //{
                            //    int reTry = 0;
                            //    while (reTry < 4)
                            //    {
                            //        Thread.Sleep(500);
                            //        LogArrayLastSaveIndex = DictGetLightStatusLog[CurThreadName].LogArrayLastSaveIndex;
                            //        LogArrayLogIndex = DictGetLightStatusLog[CurThreadName].LogArrayLogIndex;

                            //        strStatus = String.Format("{0} Status {1}: {2}", TAG, LogArrayLogIndex, DictGetLightStatusLog[CurThreadName].LogArray[LogArrayLastSaveIndex]);
                            //        Last50LogArray[LogArrayLastSaveIndex] = strStatus;
                            //        reTry++;
                            //    }
                            //}
                            #endregion


                            if (iCurLoop == 1) // 第一個Loop先不判斷資料
                            {
                                LogArrayLastSaveIndex = newLastSaveIndex;
                            }
                            else  //-- iCurLoop != 1
                            {
                                #region
                                if ((newLastSaveIndex != LogArrayLastSaveIndex) && strStatus.Contains("\"commandId\":\"0x0B\"") && (strStatus.Contains("\"0x0100\"") || strStatus.Contains("\"0x0000\"")))
                                {
                                    #region
                                    //string[] sArray = strStatus.Split(',');
                                    //foreach (string str in sArray)
                                    //{
                                    //    if (str.Contains("commandData"))
                                    //    {
                                    //        string[] dataArray = str.Split(':');
                                    //        currentResult = dataArray[1];   // "0x0100" or "0x0000"
                                    //        break;
                                    //    }
                                    //}
                                    #endregion
                                    bGetData = true;
                                    LogArrayLastSaveIndex = newLastSaveIndex;
                                }
                                else if ((newLastSaveIndex != LogArrayLastSaveIndex) && ((newLastSaveIndex) != (LogArrayLastSaveIndex + 1)))  //-- 想要的訊息不在最後一筆，設定flag後，在下面程式往回找 Data
                                {
                                    bDataIsLastOne = false;
                                    //for (int iReadIndex = (newLastSaveIndex - 1); iReadIndex >= (LogArrayLastSaveIndex + 1); iReadIndex--)
                                    //{
                                    //    newLastLogIndex--;
                                    //    strStatus = String.Format("{0} Status {1}: {2}", TAG, newLastLogIndex, DictGetLightStatusLog[CurThreadName].LogArray[iReadIndex]);
                                    //    Last50LogArray[iReadIndex] = strStatus;
                                    //    if (strStatus.Contains("\"commandId\":\"0x0B\"") && (strStatus.Contains("\"0x0100\"") || strStatus.Contains("\"0x0000\"")))
                                    //    {
                                    //        bGetData = true;
                                    //        LogArrayLastSaveIndex = newLastSaveIndex;
                                    //        break;
                                    //    }
                                    //}
                                }
                                else if (newLastSaveIndex == LogArrayLastSaveIndex)  //-- 想要的訊息尚未 Save在Dictionary，Retry
                                {
                                    int reTry = 0;
                                    while ((newLastSaveIndex == LogArrayLastSaveIndex) && reTry < 4)
                                    {
                                        Thread.Sleep(500);
                                        newLastSaveIndex = DictGetLightStatusLog[CurThreadName].LogArrayLastSaveIndex;
                                        newLastLogIndex  = DictGetLightStatusLog[CurThreadName].LogArrayLogIndex;

                                        strStatus = String.Format("{0} Status {1}: {2}", TAG, newLastLogIndex, DictGetLightStatusLog[CurThreadName].LogArray[newLastSaveIndex]);
                                        Last50LogArray[newLastSaveIndex] = strStatus;
                                        reTry++;
                                    }



                                    if ((newLastSaveIndex != LogArrayLastSaveIndex) && strStatus.Contains("\"commandId\":\"0x0B\"") && (strStatus.Contains("\"0x0100\"") || strStatus.Contains("\"0x0000\"")))
                                    {
                                        #region
                                        //string[] sArray = strStatus.Split(',');
                                        //foreach (string str in sArray)
                                        //{
                                        //    if (str.Contains("commandData"))
                                        //    {
                                        //        string[] dataArray = str.Split(':');
                                        //        currentResult = dataArray[1];   // "0x0100" or "0x0000"
                                        //        break;
                                        //    }
                                        //}
                                        #endregion
                                        bGetData = true;
                                        LogArrayLastSaveIndex = newLastSaveIndex;
                                    }
                                    else if ((newLastSaveIndex != LogArrayLastSaveIndex) && ((newLastSaveIndex) != (LogArrayLastSaveIndex + 1)))  //-- 想要的訊息不在最後一筆，設定flag後，在下面程式往回找 Data
                                    {
                                        bDataIsLastOne = false;
                                        //for (int iReadIndex = (newLastSaveIndex - 1); iReadIndex >= (LogArrayLastSaveIndex + 1); iReadIndex--)
                                        //{
                                        //    newLastLogIndex--;
                                        //    strStatus = String.Format("{0} Status {1}: {2}", TAG, newLastLogIndex, DictGetLightStatusLog[CurThreadName].LogArray[iReadIndex]);
                                        //    Last50LogArray[iReadIndex] = strStatus;
                                        //    if (strStatus.Contains("\"commandId\":\"0x0B\"") && (strStatus.Contains("\"0x0100\"") || strStatus.Contains("\"0x0000\"")))
                                        //    {
                                        //        bGetData = true;
                                        //        LogArrayLastSaveIndex = newLastSaveIndex;
                                        //        break;
                                        //    }
                                        //}
                                    }
                                }


                                if (bDataIsLastOne == false)
                                {
                                    for (int iReadIndex = (newLastSaveIndex - 1); iReadIndex >= (LogArrayLastSaveIndex + 1); iReadIndex--)
                                    {
                                        newLastLogIndex--;
                                        strStatus = String.Format("{0} Status {1}: {2}", TAG, newLastLogIndex, DictGetLightStatusLog[CurThreadName].LogArray[iReadIndex]);
                                        Last50LogArray[iReadIndex] = strStatus;
                                        if (strStatus.Contains("\"commandId\":\"0x0B\"") && (strStatus.Contains("\"0x0100\"") || strStatus.Contains("\"0x0000\"")))
                                        {
                                            bGetData = true;
                                            LogArrayLastSaveIndex = newLastSaveIndex;
                                            break;
                                        }
                                    }
                                }

                                if (bGetData == true)
                                {
                                    string[] sArray = strStatus.Split(',');
                                    foreach (string str in sArray)
                                    {
                                        if (str.Contains("commandData"))
                                        {
                                            string[] dataArray = str.Split(':');
                                            currentResult = dataArray[1];   // "0x0100" or "0x0000"
                                            break;
                                        }
                                    }
                                    LogArrayLastSaveIndex = newLastSaveIndex;
                                }
                                #endregion  //-- (iCurLoop != 1)

                                //
                                //----------------- 判斷 PASS 或 FAIL -----------------//
                                //
                                if (a != 0 && ((ListAction[a - 1].CompareTo("Power On") == 0 && currentResult.CompareTo("\"0x0100\"") == 0) ||
                                               (ListAction[a - 1].CompareTo("Power Off") == 0 && currentResult.CompareTo("\"0x0000\"") == 0)))
                                {
                                    Info = String.Format("{0} -----------------------------------> PASS", TAG);
                                    //PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                                    bResult = true;
                                }
                                else
                                {
                                    Info = String.Format("{0} -----------------------------------> FAIL", TAG);
                                    PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                                    bResult = false;
                                }
                            }

                            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { strStatus, txtPXzigbeePowerOnOffFunctionTestInformation });
                            PXzigbeePowerOnOffWriteLog.WriteLine(strStatus);
                            structNode_tmp = DictGetLightStatusLog[CurThreadName];


                            #region Old Code
                            /*
                            int reTry = 0;
                            while ( (LogArrayLastReadIndex == LogArrayLastSaveIndex) && reTry < 4)
                            {
                                Thread.Sleep(500);
                                LogArrayLastSaveIndex = DictGetLightStatusLog[CurThreadName].LogArrayLastSaveIndex;
                                LogArrayLogIndex = DictGetLightStatusLog[CurThreadName].LogArrayLogIndex;
                                
                                strStatus = String.Format("{0} Status {1}: {2}", TAG, LogArrayLogIndex, DictGetLightStatusLog[CurThreadName].LogArray[LogArrayLastSaveIndex]);
                                Last50LogArray[LogArrayLastSaveIndex] = strStatus;
                                PXzigbeePowerOnOffWriteLog.WriteLine(strStatus);
                                reTry++;
                            }
                            LogArrayLastReadIndex = LogArrayLastSaveIndex;
                            strStatus = String.Format("{0} Status {1}: {2}", TAG, LogArrayLogIndex, DictGetLightStatusLog[CurThreadName].LogArray[LogArrayLastSaveIndex]);
                            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { strStatus, txtPXzigbeePowerOnOffFunctionTestInformation });
                            PXzigbeePowerOnOffWriteLog.WriteLine(strStatus);
                            Last50LogArray[LogArrayLastSaveIndex] = strStatus;
                            structNode_tmp = DictGetLightStatusLog[CurThreadName];

                            if ( strStatus.Contains("\"commandId\":\"0x0B\"") && (strStatus.Contains("\"0x0100\"") || strStatus.Contains("\"0x0000\"")) )
                            {
                                string[] sArray = strStatus.Split(',');
                                foreach (string str in sArray)
                                {
                                    if (str.Contains("commandData"))
                                    {
                                        string[] dataArray = str.Split(':');
                                        currentResult = dataArray[1];   // "0x0100" or "0x0000"
                                        break;
                                    }
                                }

                                //if (currentResult.CompareTo(previousResult) != 0
                                if(a != 0 && ( (ListAction[a-1].CompareTo("Power On") == 0  && currentResult.CompareTo("\"0x0100\"") == 0) ||
                                               (ListAction[a-1].CompareTo("Power Off") == 0 && currentResult.CompareTo("\"0x0000\"") == 0)  ))
                                {
                                    Info = String.Format("{0} -----------------------------------> PASS", TAG);
                                    //PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                                    bResult = true;
                                    previousResult = currentResult;
                                }
                                //else if (currentResult.CompareTo(previousResult) == 0)
                                else
                                {
                                    Info = String.Format("{0} -----------------------------------> FAIL", TAG);
                                    PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                                    bResult = false;
                                }
                            }
                            */
                            #endregion


                            break;
                            #endregion

                        #region
                        /*
                        case "__Get Light Status__":
                            #region
                            Info = String.Format("{0} Get Light Status...", TAG);
                            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                            PXzigbeePowerOnOffWriteLog.WriteLine(Info);

                            PowerOnOffTestTimer.Reset();
                            PowerOnOffTestTimer.Start();
                            bTimeOut = false;

                            int iPeek = StatusProcess.StandardOutput.Peek();
                            Info = String.Format("{0} Peek: {1}", TAG, iPeek.ToString());
                            PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                            while (PowerOnOffTestTimer.ElapsedMilliseconds <= 3000 && ((tempOutput = StatusProcess.StandardOutput.ReadLine()) != null || (tempOutput = StatusProcess.StandardError.ReadLine()) != null))
                            //while (PowerOnOffTestTimer.ElapsedMilliseconds <= 3000 && StatusProcess.StandardOutput.Peek() > -1)
                            {
                                currentResult = "";
                                //tempOutput = StatusProcess.StandardOutput.ReadLine();
                                //output = tempOutput + "\n";
                                output = tempOutput;

                                Info = String.Format("{0} Status: {1}", TAG, output);
                                Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                                PXzigbeePowerOnOffWriteLog.WriteLine(Info);

                                if (output.Contains("\"0x0100\"") || output.Contains("\"0x0000\""))
                                {
                                    string[] sArray = output.Split(',');
                                    foreach (string str in sArray)
                                    {
                                        //PXzigbeePowerOnOffWriteLog.WriteLine(str);
                                        if (str.Contains("commandData"))
                                        {
                                            string[] dataArray = str.Split(':');
                                            currentResult = dataArray[1];   // "0x0100" or "0x0000"
                                            break;
                                        }
                                    }
                                }
                            
                                if ( (currentResult.CompareTo("\"0x0000\"") == 0 || currentResult.CompareTo("\"0x0100\"") == 0) && currentResult.CompareTo(previousResult) != 0)
                                {
                                    Info = String.Format("{0} -----------------------------> PASS", TAG);
                                    PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                                    bResult = true;
                                    previousResult = currentResult;
                                    break;
                                }
                                else if ((currentResult.CompareTo("\"0x0000\"") == 0 || currentResult.CompareTo("\"0x0100\"") == 0) && currentResult.CompareTo(previousResult) == 0)
                                {
                                    Info = String.Format("{0} -----------------------------> FAIL", TAG);
                                    PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                                    bResult = false;
                                    break;
                                }
                                
                            }
                            PowerOnOffTestTimer.Stop();
                            #endregion
                            break;
                        */
                        #endregion

                        default:
                            return;
                    } //--switch (strActionName)

                    Info = String.Format("{0} Wait {1} seconds...", TAG, ListSleepTimer[a]);
                    Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                    PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                    Info = String.Format("{0}", TAG);
                    Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                    PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                    Thread.Sleep(strSleepTimer);


                    if (iCurLoop != 1 && strActionName.CompareTo("Get Light Status") == 0 && bResult == false)
                    {
                        Info = String.Format("{0} ============= Test Fail !! =============", TAG, ListSleepTimer[a]);
                        Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
                        PXzigbeePowerOnOffWriteLog.WriteLine(Info);
                        bTestFail = true;
                        break;
                    }

                    #endregion //-- Execution Step Loop
                } //-- for (int a = 0; a < 4; a++)

                if (iCurLoop != 1 && structPXzigbeePowerOnOffTestSetting.StopWhenTestFailed == true && bTestFail == true)
                {
                    bPXzigbeePowerOnOffFTThreadRunning = false;
                    break;
                }

                #endregion //-- Main Test Loop
            }
            //PowerOnProcess.WaitForExit();
            //PowerOnProcess.Kill();
            //PowerOnProcess.Close();

            //-- Write Last 50 Light Status Log --//
            PXzigbeePowerOnOffWriteLog.WriteLine(TAG);
            PXzigbeePowerOnOffWriteLog.WriteLine(TAG + " Last 50 Light Status Log:");
            for (int j = 0; j < Last50LogArray.Length; j++)
            {
                PXzigbeePowerOnOffWriteLog.WriteLine(Last50LogArray[j]);
            }
            structNode_tmp.writeLogFinished = true;
            DictGetLightStatusLog[CurThreadName] = structNode_tmp;
            PXzigbeePowerOnOffWriteLog.Close();
            

            //
            //-- Thread test finish, count +1 --//
            //
            iTestFinishCount = iTestFinishCount + 1;
        }

        private void DoPXzigbeePowerOnOffGetLightStatus(object CurThreadName)
        {
            int iCurIndex = 0;
            //string Info = "";
            string tempOutput = "";
            string outputforLog = "";
            string CUR_THREAD_NAME = CurThreadName.ToString();

            PXzigbeePowerOnOff_GetLightStatusLog structNode_tmp = new PXzigbeePowerOnOff_GetLightStatusLog();
            structNode_tmp.LogArray = new string[50];      // Array
            structNode_tmp.LogList  = new List<string>();  // List

            string DebugFilePath = System.Windows.Forms.Application.StartupPath + @"\DebugMessage.txt";
            System.IO.StreamWriter WriteDebugMsg;
            WriteDebugMsg = new System.IO.StreamWriter(DebugFilePath);
            WriteDebugMsg.AutoFlush = true;

            
            structNode_tmp.LogList.Add("First Log");
            DictGetLightStatusLog.Add(CUR_THREAD_NAME, structNode_tmp);
            DictGetLightStatusLog[CUR_THREAD_NAME].LogList.Clear();
            

            //-- Find structure index for current thread (current test item) --//
            for (int iIndex = 0; iIndex < iTestItemsCount; iIndex++)
            {
                if (structPXzigbeePowerOnOffConfig[iIndex].ThreadName.CompareTo(CUR_THREAD_NAME) == 0)
                {
                    iCurIndex = iIndex;
                    break;
                }
            }

            string ip = structPXzigbeePowerOnOffConfig[iCurIndex].IP;
            string id = structPXzigbeePowerOnOffConfig[iCurIndex].NodeID;
            string MQTT_Path = structPXzigbeePowerOnOffConfig[iCurIndex].MQTT_Path;  //-- ex:C:\Program Files\Project X\mos158\
            string TAG = String.Format(" {0} | {1} |", CUR_THREAD_NAME, id);

            
                                           // mosquitto_pub -h 192.168.18.1 -t "gw/commands" -q 2 -m "{\"commands\":[{\"commandcli\":\"zcl global read 0x0006 0x0000\"},{\"commandcli\":\"send 0x1F57 1 0xff\"}]}"
            //string StatusCmd = String.Format("mosquitto_pub -h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl global read 0x0006 0x0000\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xff\\\"}}]}}\"", ip, id);
            string StatusCmd = String.Format("-h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl global read 0x0006 0x0000\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xff\\\"}}]}}\"", ip, id);
            //string MQTT_PathCmd = String.Format("cd {0}", MQTT_Path);                                        // cd C:\Program Files\Project X\mos158
            //string MQTT_responseCmd = String.Format("mosquitto_sub -h {0} -t \"gw/zclresponse\" -q 2", ip);  // mosquitto_sub -h 192.168.18.1 -t \"gw/zclresponse\" -q 2
            string MQTT_responseCmd = String.Format("-h {0} -t \"gw/zclresponse\" -q 2", ip);  // mosquitto_sub -h 192.168.18.1 -t \"gw/zclresponse\" -q 2


            Process StatusProcess = new Process();
            ProcessStartInfo StatusStartInfo = new ProcessStartInfo();
            StatusStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            StatusStartInfo.CreateNoWindow = true;
            StatusStartInfo.FileName = MQTT_Path + "\\mosquitto_sub.exe";
            StatusStartInfo.Arguments = MQTT_responseCmd;
            StatusStartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
            StatusStartInfo.UseShellExecute = false;
            StatusStartInfo.RedirectStandardOutput = true;
            StatusStartInfo.RedirectStandardError = true;
            StatusStartInfo.RedirectStandardInput = true;
            StatusProcess.StartInfo = StatusStartInfo;

            //StatusProcess.StandardInput.WriteLine("C:");
            //StatusProcess.StandardInput.WriteLine(MQTT_PathCmd);
            //StatusProcess.StandardInput.WriteLine(StatusCmd);
            //StatusProcess.StandardInput.WriteLine(MQTT_responseCmd);

            string startFileName = String.Format("\"{0}\\mosquitto_pub.exe\"", MQTT_Path);
            //Process.Start("\"C:\\Program Files\\Project X\\mos158\\mosquitto_pub.exe\"", StatusCmd);
            Process.Start(startFileName, StatusCmd);
            StatusProcess.Start();

            int i = 0;
            int logIndex = 0;
            while (bPXzigbeePowerOnOffFTThreadRunning == true)
            {
                //
                //-- Save data in List
                //
                #region
                /*
                if (DictGetLightStatusLog[CUR_THREAD_NAME].LogList.Count == 50)
                    DictGetLightStatusLog[CUR_THREAD_NAME].LogList.Clear();

                tempOutput = StatusProcess.StandardOutput.ReadLine();
                //outputforLog = tempOutput + "\n";
                outputforLog = tempOutput;

                Info = String.Format("{0} Status {1}: {2}", TAG, i, outputforLog);
                DictGetLightStatusLog[CUR_THREAD_NAME].LogList.Add(Info);
                i++;
                */
                #endregion

                //
                //-- Save data in Array
                //
                tempOutput = StatusProcess.StandardOutput.ReadLine();
                outputforLog = tempOutput;

                structNode_tmp.LogArray[i]           = outputforLog;
                structNode_tmp.LogArrayLastSaveIndex = i;
                structNode_tmp.LogArrayLogIndex      = logIndex;
                DictGetLightStatusLog[CUR_THREAD_NAME] = structNode_tmp;
                WriteDebugMsg.WriteLine(logIndex.ToString() + ":  " + outputforLog);
                
                if (i == 49)
                    i = 0;
                i++;
                logIndex++;

            }
            WriteDebugMsg.Close();
            
        }

        private void TestThreadCount(string CurThreadName)
        {
            int j = 0;
            string Info = "";

            j = j + 1;
            Info = "TestCount    | Current Thread: " + CurThreadName;
            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
            Info = "TestCount    | j: " + j;
            Invoke(new PXzigbeePowerOnOffSetTextCallBack(SetText), new object[] { Info, txtPXzigbeePowerOnOffFunctionTestInformation });
        }

        #endregion //-- Main Test Structure


        //=================================================================//
        //----------------------- Save Test Parameter ---------------------//
        //=================================================================//
        #region Save Test Parameter
        private bool SaveTestConfig_PXzigbeePowerOnOff()
        {
            iTestItemsCount = dgvPXzigbeePowerOnOffTestConditionData.RowCount - 1;
            string[,] rowdata = new string[iTestItemsCount, dgvPXzigbeePowerOnOffTestConditionData.ColumnCount];
            structPXzigbeePowerOnOffConfig = new PXzigbeePowerOnOff_TestConfig[iTestItemsCount];

            for (int i = 0; i < dgvPXzigbeePowerOnOffTestConditionData.RowCount - 1; i++)
            {
                for (int j = 0; j < dgvPXzigbeePowerOnOffTestConditionData.ColumnCount; j++)
                {
                    rowdata[i, j] = dgvPXzigbeePowerOnOffTestConditionData.Rows[i].Cells[j].Value.ToString();
                }
            }
            testConfigPXzigbeePowerOnOff = rowdata;

            for (int i = 0; i < iTestItemsCount; i++)
            {
                structPXzigbeePowerOnOffConfig[i].ThreadName = "";
                structPXzigbeePowerOnOffConfig[i].ModeName    = testConfigPXzigbeePowerOnOff[i, 0];
                structPXzigbeePowerOnOffConfig[i].IP          = testConfigPXzigbeePowerOnOff[i, 1];
                structPXzigbeePowerOnOffConfig[i].NodeID      = testConfigPXzigbeePowerOnOff[i, 2];
                structPXzigbeePowerOnOffConfig[i].Action1     = testConfigPXzigbeePowerOnOff[i, 3];
                structPXzigbeePowerOnOffConfig[i].SleepTimer1 = testConfigPXzigbeePowerOnOff[i, 4];
                structPXzigbeePowerOnOffConfig[i].Action2     = testConfigPXzigbeePowerOnOff[i, 5];
                structPXzigbeePowerOnOffConfig[i].SleepTimer2 = testConfigPXzigbeePowerOnOff[i, 6];
                structPXzigbeePowerOnOffConfig[i].Action3     = testConfigPXzigbeePowerOnOff[i, 7];
                structPXzigbeePowerOnOffConfig[i].SleepTimer3 = testConfigPXzigbeePowerOnOff[i, 8];
                structPXzigbeePowerOnOffConfig[i].Action4     = testConfigPXzigbeePowerOnOff[i, 9];
                structPXzigbeePowerOnOffConfig[i].SleepTimer4 = testConfigPXzigbeePowerOnOff[i, 10];
                structPXzigbeePowerOnOffConfig[i].MQTT_Path   = testConfigPXzigbeePowerOnOff[i, 11];
            }
            return true;
        }

        private void PXzigbeePowerOnOffSaveTestSetting()
        {
            structPXzigbeePowerOnOffTestSetting.TestName = txtPXzigbeePowerOnOffFunctionTestName.Text;
            structPXzigbeePowerOnOffTestSetting.TestTimes = nudPXzigbeePowerOnOffFunctionTestTimes.Value.ToString();
            structPXzigbeePowerOnOffTestSetting.StopWhenTestFailed = cboxPXzigbeePowerOnOffFunctionTestStopWhenTestFailed.Checked;
        }
        #endregion //-- Save Test Parameter


        //=================================================================//
        //------------------------ Invoke Function ------------------------//
        //=================================================================//
        #region Invoke Function

        //-----------------  System Cursor ------------------//
        #region System Cursor
        private void TogglePXzigbeePowerOnOffFunctionTestSystemCursorWait()
        {
            TogglePXzigbeePowerOnOffFunctionTestSystemCursorStatus(true);
        }

        private void TogglePXzigbeePowerOnOffFunctionTestSystemCursorDefault()
        {
            TogglePXzigbeePowerOnOffFunctionTestSystemCursorStatus(false);
        }

        private void TogglePXzigbeePowerOnOffFunctionTestSystemCursorStatus(bool Toggle)
        {
            if (Toggle == true)
            {
                btnPXzigbeePowerOnOffFunctionTestRun.Enabled = false;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            }
            else
            {
                btnPXzigbeePowerOnOffFunctionTestRun.Enabled = true;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
        }

        #endregion //-- System Cursor
        //---------------------------------------------------//


        //------------------ Test Controller ----------------//
        #region Test Controller
        private void TogglePXzigbeePowerOnOffFunctionTestSettingTrue()
        {
            TogglePXzigbeePowerOnOffFunctionTestController(true);
        }

        private void TogglePXzigbeePowerOnOffFunctionTestController(bool Toggle)
        {
            btnPXzigbeePowerOnOffFunctionTestRun.Text = Toggle ? "Run" : "Stop";


            //-------- MenuItem -------//
            fileToolStripMenuItem.Enabled = Toggle;
            itemToolStripMenuItem.Enabled = Toggle;
            setupToolStripMenuItem.Enabled = Toggle;
            helpToolStripMenuItem.Enabled = Toggle;

            //----- Function Test -----//
            gbox_PXzigbeePowerOnOffFunctionTestSetting.Enabled = Toggle;

            //---- Test Condition -----//
            gbox_PXzigbeePowerOnOffTestConditionDeviceSetting.Enabled = Toggle;
            btnPXzigbeePowerOnOffTestConditionAddSetting.Enabled = Toggle;
            btnPXzigbeePowerOnOffTestConditionEditSetting.Enabled = Toggle;
            btnPXzigbeePowerOnOffTestConditionSaveSetting.Enabled = Toggle;
            btnPXzigbeePowerOnOffTestConditionLoadSetting.Enabled = Toggle;
            dgvPXzigbeePowerOnOffTestConditionData.Enabled = Toggle;


            /* Show Total Elasped Eesting Time */
            if (!Toggle)
            {
                //System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                swElapsedTime.Restart();
                timerElaspedTime.Enabled = true;
                timerElaspedTime.Start();
            }
            else
            {
                //System.Windows.Forms.Cursor.Current = Cursors.Default;
                swElapsedTime.Stop();
                timerElaspedTime.Stop();
                timerElaspedTime.Enabled = false;
            }


            if (bPXzigbeePowerOnOffFTThreadRunning == false)
            {
                btnPXzigbeePowerOnOffFunctionTestRun.Enabled = false;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

                try
                {
                    //** Remove the "mosquitto_sub.exe" process
                    System.Diagnostics.Process[] mqttProc = System.Diagnostics.Process.GetProcessesByName("mosquitto_sub");
                    for (int i = 0; i < mqttProc.Length; i++)
                        mqttProc[i].Kill();
                }
                catch { }

                try
                {
                    //** Remove the "mosquitto_pub.exe" process
                    System.Diagnostics.Process[] mqtt_pubProc = System.Diagnostics.Process.GetProcessesByName("mosquitto_pub");
                    for (int i = 0; i < mqtt_pubProc.Length; i++)
                        mqtt_pubProc[i].Kill();
                }
                catch { }

                try
                {
                    //** Remove the "cmd.exe" process
                    System.Diagnostics.Process[] cmdProc = System.Diagnostics.Process.GetProcessesByName("cmd");
                    for (int i = 0; i < cmdProc.Length; i++)
                        cmdProc[i].Kill();
                }
                catch { }

                try
                {
                    if (threadPXzigbeePowerOnOffGetStatus.IsAlive)
                    {
                        threadPXzigbeePowerOnOffGetStatus.Abort();
                        threadPXzigbeePowerOnOffGetStatus.Join();
                    }
                }
                catch { }

                try
                {
                    if (threadPXzigbeePowerOnOffFTstopEvent.IsAlive)
                    {
                        threadPXzigbeePowerOnOffFTstopEvent.Abort();
                        threadPXzigbeePowerOnOffFTstopEvent.Join();
                    }
                }
                catch { }

                for (int iIndex = 0; iIndex < iTestItemsCount; iIndex++)
                {
                    int reTry = 0;
                    string CurThreadName = "Node-" + (iIndex + 1).ToString();
                    while(DictGetLightStatusLog[CurThreadName].writeLogFinished == false && reTry < 5)
                    {
                        Thread.Sleep(1000);
                        reTry++;
                    }
                }

                try
                {
                    if (threadPXzigbeePowerOnOffMainFunction.IsAlive)
                    {
                        threadPXzigbeePowerOnOffMainFunction.Abort();
                        threadPXzigbeePowerOnOffMainFunction.Join();
                    }
                }
                catch { }

                try
                {
                    if (threadPXzigbeePowerOnOffFT.IsAlive)
                    {
                        threadPXzigbeePowerOnOffFT.Abort();
                        threadPXzigbeePowerOnOffFT.Join();
                    }
                }
                catch { }

                try
                {
                    TestFinishedActionPXzigbeePowerOnOff();
                }
                catch { }

                MessageBoxTopMost("ATE Information", "Test complete !!!");  // 將 MessageBox於桌面置頂
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                btnPXzigbeePowerOnOffFunctionTestRun.Enabled = true;
                bPXzigbeePowerOnOffTestComplete = true;
            }
        }
        #endregion //-- Test Controller
        //---------------------------------------------------//

        #endregion //-- Invoke Function


#endregion
    } //End of class 
}