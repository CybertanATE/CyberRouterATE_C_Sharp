﻿///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RouterTestMain.cs
///  Update         : 2016-07-31    
///  Version        : 1.0.160731
///  Description    : Changes to this file may cause incorrect behavior and will be lost
///                   if the code is regenerated.
///  Modified       : 2016-07-25 Initial version
///                   2016-07-25
///                     1. Add RvR Test Function
///                     2. Add Throughput Test Function
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
using System.Reflection;
using Microsoft.Win32;
using System.Threading;
using System.IO;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        Stopwatch RouterTestTimer;
        TextBox TEXTBOX_INFO = null;
        string s_RouterReportPath = System.Windows.Forms.Application.StartupPath + @"\report";

        FinalTestItemsScriptData_RouterTest[] st_ReadTestItemsScript_RouterTest;
        StepsScriptData_RouterTest[] st_ReadStepsScriptData_RouterTest;
        DeviceWebGuiScriptData_RouterTest[] st_ReadDeviceWebGuiScriptData_RouterTest;

        int TEST_ITEMS_SCRIPT_ST_IDX_RouterTest = 0;
        int STEP_SCRIPT_ST_IDX_RouterTest = 0;
        int DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest = 0;
        int i_TestRun_RouterTest = 1;


        /* Declare public member data */
        Stopwatch swElapsedTime = new Stopwatch();
        TimeSpan tsElapsedTime;
        string strElapsedTime;
        string sTestItem = string.Empty; //Identify whick test item is launched. Use on "save config" and "load config" ; 
        string sDebugFilePath = System.Windows.Forms.Application.StartupPath + @"\DebugMessage.txt";
        string sConsoleLogFilePath = System.Windows.Forms.Application.StartupPath + @"\ConsoleLog.txt";
        Dictionary<int, string> dic_LineAuth = new Dictionary<int, string>();

        //string startupPath = System.Windows.Forms.Application.StartupPath;

        public RouterTestMain()
        {
            InitializeComponent();

            #region Test Code
            //--- Write File Example ---//
            //string DebugFilePath = System.Windows.Forms.Application.StartupPath + @"\DebugMessage.txt";
            //System.IO.StreamWriter WriteDebugMsg;
            //WriteDebugMsg = new System.IO.StreamWriter(DebugFilePath);
            //WriteDebugMsg.AutoFlush = true;
            //WriteDebugMsg.Write("FFFFFFFFFFFFFFFF");  //Write()     不換行
            //WriteDebugMsg.WriteLine("       TTTT");   //WriteLine() 會換行
            //WriteDebugMsg.WriteLine("Test12345");
            //WriteDebugMsg.Close();



            //--- Read/Write Cmd ---//
            //CmdTest();

            //ProcessTest();
            //CmdProcessTest();
            //DictTest();


            /*
            string ip = "192.168.18.1";
            string id = "0xAB24";
            string PowerOnCmd  = String.Format("-h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl on-off on\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xFF\\\"}}]}}\"", ip, id);
            string PowerOffCmd = String.Format("-h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl on-off off\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xFF\\\"}}]}}\"", ip, id);
            
            //Process.Start("\"C:\\Program Files\\Project X\\mos158\\mosquitto_pub.exe\"", PowerOnCmd);
            //Thread.Sleep(3000);
            //Process.Start("\"C:\\Program Files\\Project X\\mos158\\mosquitto_pub.exe\"", PowerOffCmd);
            

            
            Process PowerOnProcess = new Process();
            ProcessStartInfo PowerOnstartInfo = new ProcessStartInfo();
            PowerOnstartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            PowerOnstartInfo.CreateNoWindow = true;
            PowerOnstartInfo.FileName = "\"C:\\Program Files\\Project X\\mos158\\mosquitto_pub.exe\"";
            PowerOnstartInfo.Arguments = PowerOnCmd;
            PowerOnstartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
            PowerOnstartInfo.UseShellExecute = false;
            PowerOnstartInfo.RedirectStandardOutput = true;
            PowerOnstartInfo.RedirectStandardError = true;
            PowerOnstartInfo.RedirectStandardInput = true;
            PowerOnProcess.StartInfo = PowerOnstartInfo;


            Process PowerOffProcess = new Process();
            ProcessStartInfo PowerOffstartInfo = new ProcessStartInfo();
            PowerOffstartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            PowerOffstartInfo.CreateNoWindow = true;
            PowerOffstartInfo.FileName = "\"C:\\Program Files\\Project X\\mos158\\mosquitto_pub.exe\"";
            PowerOffstartInfo.Arguments = PowerOffCmd;
            PowerOffstartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
            PowerOffstartInfo.UseShellExecute = false;
            PowerOffstartInfo.RedirectStandardOutput = true;
            PowerOffstartInfo.RedirectStandardError = true;
            PowerOffstartInfo.RedirectStandardInput = true;
            PowerOffProcess.StartInfo = PowerOffstartInfo;


            PowerOnProcess.Start();
            PowerOnProcess.WaitForExit();
            PowerOnProcess.Close();
            Thread.Sleep(2000);

            PowerOffProcess.Start();
            PowerOffProcess.WaitForExit();
            PowerOffProcess.Close();
            Thread.Sleep(2000);

            PowerOnProcess.Start();
            PowerOnProcess.WaitForExit();
            PowerOnProcess.Close();
            Thread.Sleep(2000);

            PowerOffProcess.Start();
            PowerOffProcess.WaitForExit();
            PowerOffProcess.Close();
            */

            //Process.Start("ipconfig");
            #endregion
        }


        /// <summary>
        /// OnLoad event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param
        private void OnLoad(object sender, EventArgs e)
        {
            /* Check .net version */
            if (!CheckDotNetVersion())
            {
                MessageBox.Show("This program needs to setup .Net Framework 4.0 above.", "Error");
                Application.Exit();
            }

            /* Displays copyright and version number in titlebar */
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            this.Text = String.Format("CyberRouterATE v{0} - Copyright(c) 2014-2020 CyberTan", fvi.FileVersion);

            /* Pre-create subfolder if report and config folders are not exist */
            PreCreateSubFolder();
            ToggleToolStripMenuItem(false);
            tabControl_RouterStartPage.Show();
            tsslMessage.Text = tabControl_RouterStartPage.TabPages[tabControl_RouterStartPage.SelectedIndex].Text + " Control Panel";
        }

        /// <summary>
        /// Form closing event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RouterTestMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Do you want to exit the program?", "Exit", MessageBoxButtons.YesNo) == DialogResult.Yes)
                CloseAllResources();
            else
                e.Cancel = true;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutCyberRouterATE ate = new AboutCyberRouterATE();
            ate.ShowDialog();    
        }

        private void serialPortToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ConfigSerialPort serialPort = new ConfigSerialPort();            
            serialPort.ShowDialog();
        }

        private void serialPort2ToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            ConfigSerialPort2 serialPort2 = new ConfigSerialPort2();
            serialPort2.ShowDialog();
        }

        private void lineNotifyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ConfigLineNotify lineNotify = new ConfigLineNotify();
            lineNotify.ShowDialog();
        }
        
        private void dutControlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DutControll dutControl = new DutControll();
            dutControl.ShowDialog();            
        }

        private void timerElaspedTime_Tick(object sender, EventArgs e)
        {
            // Get the elapsed time as a TimeSpan value.
            tsElapsedTime = swElapsedTime.Elapsed;

            // Format and display the TimeSpan value.
            strElapsedTime = String.Format("{0:00}-{1:00}:{2:00}:{3:00}",
                tsElapsedTime.Days,
                tsElapsedTime.Hours, tsElapsedTime.Minutes, tsElapsedTime.Seconds);

            labElapsedTime.Text = "Elapsed time: " + strElapsedTime;
        }
        
        /// <summary>
        /// Pre-create config and report subfolders
        /// </summary>
        private void PreCreateSubFolder()
        {
            string subPath = "\\config";

            subPath = System.Windows.Forms.Application.StartupPath + subPath;

            bool isExists = System.IO.Directory.Exists(subPath);

            if (!isExists)
                System.IO.Directory.CreateDirectory(subPath);

            subPath = System.Windows.Forms.Application.StartupPath + "\\report";
            isExists = System.IO.Directory.Exists(subPath);
            if (!isExists)
                System.IO.Directory.CreateDirectory(subPath);

            subPath = System.Windows.Forms.Application.StartupPath + "\\testCondition";
            isExists = System.IO.Directory.Exists(subPath);
            if (!isExists)
                System.IO.Directory.CreateDirectory(subPath);
        }

        /// <summary>
        /// Save the Information into a log file.
        /// </summary>
        /// <returns></returns>
        private void SaveInofrmationText(string savePath, string str)
        {
            string Filename = DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
            string Path = savePath + "\\" + Filename;

            try
            {
                File.WriteAllText(Path, str);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Error: Total of file path is exceeds 218 characters.");
            }
        }

        /// <summary>
        /// Save the Information into a log file.
        /// </summary>
        /// <returns></returns>
        private void SaveInofrmationText(string savePath, TextBox textbox)
        {
            string Filename = DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
            string Path = savePath + "\\" + Filename;

            try
            {
                File.WriteAllText(Path, textbox.Text);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Error: Total of file path is exceeds 218 characters.");
            }
        }

        /// <summary>
        /// Check .Net framework version
        /// </summary>
        /// <returns></returns>
        private bool CheckDotNetVersion()
        {
            string componentsKeyName = "SOFTWARE\\Microsoft\\Active Setup\\Installed Components",
                   friendlyName,
                   version;
            // Find out in the registry anything under:
            //    HKLM\SOFTWARE\Microsoft\Active Setup\Installed Components
            // that has ".NET Framework" in the name
            RegistryKey componentsKey = Registry.LocalMachine.OpenSubKey(componentsKeyName);
            string[] instComps = componentsKey.GetSubKeyNames();
            foreach (string instComp in instComps)
            {
                RegistryKey key = componentsKey.OpenSubKey(instComp);
                friendlyName = (string)key.GetValue(null); // Gets the (Default) value from this key
                if (friendlyName != null && friendlyName.IndexOf(".NET Framework") >= 0)
                {
                    // Try to get any version information that's available
                    version = (string)key.GetValue("Version");
                    // Just only checked the framework information
                    string[] str = version.Split(',');
                    if (version != null && Convert.ToDouble(str[0]) >= 4)
                    {
                        Debug.WriteLine(str[0]);
                        return true;
                    }
                }
            }
            return false;
        }


        private void btnRvRFunctionTestSaveLog_Click(object sender, EventArgs e)
        {

        }





        //====================================================================//
        //--------------------------  Test Area ------------------------------//
        //====================================================================//
        #region Test Area

        private void CmdTest()
        {
            //string RunCheckSSDbatPath  = startupPath + @"\RunCheckSSDbat.bat";
            string RunWriteFileBatPath = startupPath + @"\RunWriteFileBat.bat";
            string ThirdpartyToolsPath = startupPath + "\\ThirdpartyTools";
            string PSToolsPath = ThirdpartyToolsPath + "\\PSTools";
            string BatFilePath = ThirdpartyToolsPath + "\\WriteFile.bat";


            //System.IO.StreamWriter WriteBatFile1;
            //WriteBatFile1 = new System.IO.StreamWriter(RunCheckSSDbatPath);

            //WriteBatFile2.WriteLine(RemoteCmd);
            //WriteBatFile2.Close();

            // psexec \\172.16.20.20 -u "Aspire E15" -p 123 -c -f C:\tmpFile\shutdown.bat
            //string RemoteCmd = "psexec \\\\" + Parameter1 + " -u \"user\" -p \"user\" -c -f " + SystemSleepFile;
            //string RemoteCmd = "psexec \\\\" + Parameter1 + " -u \"" + P1_LoginID + "\" -p \"" + P1_LoginPW + "\" -c -f " + SystemSleepFile;
            string RemoteCmd = "psexec \\\\192.168.18.37 -u \"user\" -p \"user\" -c -f " + BatFilePath;
            string PSpathCmd = "Cd " + PSToolsPath;

            System.IO.StreamWriter WriteBatFile2;
            WriteBatFile2 = new System.IO.StreamWriter(RunWriteFileBatPath);

            WriteBatFile2.WriteLine(PSpathCmd);
            WriteBatFile2.WriteLine(RemoteCmd);
            WriteBatFile2.Close();

            Process.Start(RunWriteFileBatPath);
        }


        struct GetLightStatusLog
        {

            public List<string> LogList;
            //public int MyInt;

            //public GetLightStatusLog(int myInt)
            //{
            //    MyInt = myInt;
            //    LogList = new List<string>();
            //}
        }

        private void DictTest()
        {
            Dictionary<string, GetLightStatusLog> DictGetLightStatus = new Dictionary<string, GetLightStatusLog>();
            GetLightStatusLog structGetLightStatusLog = new GetLightStatusLog();
            structGetLightStatusLog.LogList = new List<string>();

            string DebugFilePath = System.Windows.Forms.Application.StartupPath + @"\DebugMessage.txt";
            System.IO.StreamWriter WriteDebugMsg;
            WriteDebugMsg = new System.IO.StreamWriter(DebugFilePath);
            WriteDebugMsg.AutoFlush = true;

            structGetLightStatusLog.LogList.Add("First Log Test");

            DictGetLightStatus.Add("Node-1", structGetLightStatusLog);
            
            DictGetLightStatus["Node-1"].LogList.Add("2222222222");
            DictGetLightStatus["Node-1"].LogList.Add("3333333333333");

            WriteDebugMsg.WriteLine("LogList[0]:" + DictGetLightStatus["Node-1"].LogList[0]);

            int LogListCount = DictGetLightStatus["Node-1"].LogList.Count;
            WriteDebugMsg.WriteLine("DictGetLightStatus[\"Node-1\"].LogList.Count: " + LogListCount);


            WriteDebugMsg.WriteLine("\n\nAll LogList:\n");
            for (int i = 0; i < LogListCount; i++)
            {
                WriteDebugMsg.WriteLine(DictGetLightStatus["Node-1"].LogList[i]);
            }

            DictGetLightStatus["Node-1"].LogList.Clear();  //-------- Clear()
            DictGetLightStatus["Node-1"].LogList.Add("New Data1");
            DictGetLightStatus["Node-1"].LogList.Add("New Data2");
            LogListCount = DictGetLightStatus["Node-1"].LogList.Count;
            WriteDebugMsg.WriteLine("\n\nNew All LogList:\n");
            for (int i = 0; i < LogListCount; i++)
            {
                WriteDebugMsg.WriteLine(DictGetLightStatus["Node-1"].LogList[i]);
            }
            
            
            
            
            WriteDebugMsg.Close();
        

        
        }

        private bool CmdProcessTest()
        {
            //mosquitto_pub -h 192.168.18.1 -t "gw/commands" -q 2 -m "{\"commands\":[{\"commandcli\":\"zcl global read 0x0006 0x0000\"},{\"commandcli\":\"send 0x1F57 1 0xff\"}]}"
            string ip = "192.168.18.1";
            string id = "0xAB24";
            string cmd1 = String.Format("mosquitto_pub -h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl global read 0x0006 0x0000\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xff\\\"}}]}}\"", ip, id);

            Process StatusProcess = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = @"C:\Windows\System32" + "\\cmd.exe";
            startInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;

            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            StatusProcess.StartInfo = startInfo;
            StatusProcess.Start();

            string onCmd = String.Format("mosquitto_pub -h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl on-off on\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xFF\\\"}}]}}\"", ip, id);
            string offCmd = String.Format("mosquitto_pub -h {0} -t \"gw/commands\" -q 2 -m \"{{\\\"commands\\\":[{{\\\"commandcli\\\":\\\"zcl on-off off\\\"}},{{\\\"commandcli\\\":\\\"send {1} 1 0xFF\\\"}}]}}\"", ip, id);


            StatusProcess.StandardInput.WriteLine("C:");
            StatusProcess.StandardInput.WriteLine(@"cd C:\Program Files\Project X\mos158");
            //StatusProcess.StandardInput.WriteLine(cmd1);
            //StatusProcess.StandardInput.WriteLine("mosquitto_sub -h 192.168.18.1 -t \"gw/zclresponse\" -q 2");
            
            StatusProcess.StandardInput.WriteLine(onCmd);
            StatusProcess.StandardInput.WriteLine("timeout /t 5");
            StatusProcess.StandardInput.WriteLine(offCmd);

            string tempOutput = "";
            string output = "";

            string FilePath = startupPath + @"\GetStatusLog.txt";
            System.IO.StreamWriter WriteLog;
            WriteLog = new System.IO.StreamWriter(FilePath);


            while ((tempOutput = StatusProcess.StandardOutput.ReadLine()) != null || (tempOutput = StatusProcess.StandardError.ReadLine()) != null)
            {
                //output += tempOutput + "\n";
                output = tempOutput + "\n";
                WriteLog.Flush();
                WriteLog.WriteLine(output);
                //if (output.Contains("gw/zclresponse"))
                //    break;
            }
            WriteLog.Close();

            
            /*
            string expectedOutput = @"""status"":200,""message"":""ok""";
            if (output.Contains(expectedOutput) == false)
            {
                StatusProcess.WaitForExit();
                StatusProcess.Close();
                return false;
            }*/

            ////StatusProcess.WaitForExit();
            ////StatusProcess.Close();
            return true;
        }   
        
        private bool ProcessTest()
        {
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = @"C:\Windows\System32" + "\\cmd.exe";
            startInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;

            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            process.StartInfo = startInfo;
            process.Start();
            process.StandardInput.WriteLine("ipconfig /all");
            //process.StandardInput.WriteLine(@"curl -X POST -H ""Authorization:Bearer " + _authorization + @""" -F ""message=" + message + @""" " + _apiAddress);
            process.StandardInput.WriteLine("exit");
     
            
            string tempOutput = "";
            string output = "";


            while ((tempOutput = process.StandardOutput.ReadLine()) != null || (tempOutput = process.StandardError.ReadLine()) != null)
            {
                output += tempOutput + "\n";
            }


            string FilePath = startupPath + @"\Log.txt";
            System.IO.StreamWriter WriteLog;
            WriteLog = new System.IO.StreamWriter(FilePath);
            WriteLog.WriteLine(output);
            WriteLog.Close();
            /*
            string expectedOutput = @"""status"":200,""message"":""ok""";
            if (output.Contains(expectedOutput) == false)
            {
                process.WaitForExit();
                process.Close();
                return false;
            }*/

            process.WaitForExit();
            process.Close();
            return true;
        }
        
        #endregion //--------------------------------------------- Test Area




        /*============================================================================================*/
        /*====================== From now on, Function need to update ================================*/
        /*============================================================================================*/

        

        /// <summary>
        /// Turn all Items unCheck and Hide all the tabControl
        /// </summary>
        /// <param name="Toggle"></param>
        private void ToggleToolStripMenuItem(bool Toggle)
        {
            rvRTestToolStripMenuItem.Checked = false;
            powerOnOffTestToolStripMenuItem.Checked = false;
            rvRTurnTableTestToolStripMenuItem.Checked = false;
            interOperabilityToolStripMenuItem.Checked = false;
            throughputTestToolStripMenuItem.Checked = false;
            webGUITestToolStripMenuItem.Checked = false;
            USBStorageTestToolStripMenuItem.Checked = false;
            fWUpgradeDowngradeTestToolStripMenuItem.Checked = false;
            gUITestToolStripMenuItem.Checked = false;
			GuestNetworkTestToolStripMenuItem.Checked = false;
            chamberPerformanceTestToolStripMenuItem.Checked = false;
            integrationTestToolStripMenuItem.Checked = false;
            pxZigbeePowerOnOffTestToolStripMenuItem.Checked = false;

            RvRTest_tabControl.Hide();
            PowerOnOff_tabControl.Hide();
            PXzigbeePowerOnOff_tabControl.Hide();
            tmp_tabControl.Hide();
            tabControl_Interoperability.Hide();
            tabControl_Throughput.Hide();
            tabControl_GUItest.Hide();
            tabControl_USBStorage.Hide();
            tabControl_GUI.Hide();
			tabControl_GuestNetworkTest.Hide();
            tabControl_ChamberPerformance.Hide();
            tabControl_RouterIntegration.Hide();
        }
              
        private void saveConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filename = string.Empty;

            switch (sTestItem)
            {
                case "RvR Test":
                    filename = Constants.XML_ROUTER_RvR;
                    break;
                case "RvR Turn Test":
                    filename = Constants.XML_ROUTER_RvRTURN;
                    break;
                case "Power On Off Test":
                    filename = Constants.XML_ROUTER_POWERONOFF;
                    break;
                case "Interoperability Test":
                    filename = Constants.XML_ROUTER_INTEROPERABILITY;
                    break;
                case Constants.TESTITEM_ROUTER_ChamberPerformance:
                    filename = Constants.XML_ROUTER_ChamberPerformance;
                    break;
                default:
                    return;
            }

            // Displays a SaveFileDialog so the user can save the XML assigned to Save config.
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\config\";
            saveFileDialog1.FileName = filename;
            saveFileDialog1.DefaultExt = ".xml";
            saveFileDialog1.Filter = "XML file|*.xml";
            saveFileDialog1.Title = "Save an xml file";

            // If the file name is not an empty string open it for saving.
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName != "")
            {
                switch (sTestItem)
                {
                    case "RvR Test":
                        WriteXmlRvRTest(saveFileDialog1.FileName);
                        break;
                    case "RvR Turn Test":
                        WriteXmlRvRTurnTest(saveFileDialog1.FileName);
                        break;
                    case "Power On Off Test":
                        WriteXmlPowerOnOff(saveFileDialog1.FileName);
                        break;
                    case "Interoperability Test":
                        //WriteXmlInteroperability(saveFileDialog1.FileName);
                        break;
                    case Constants.TESTITEM_ROUTER_ChamberPerformance:
                        writeXmlRouterChamberPerformanceTest(saveFileDialog1.FileName);
                        break;
                    default:
                        return;
                }
            }
        }

        private void loadConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filename = string.Empty;

            switch (sTestItem)
            {
                
                case "RvR Test":
                    filename = Constants.XML_ROUTER_RvR;
                    break;
                case "RvR Turn Test":
                    filename = Constants.XML_ROUTER_RvRTURN;
                    break;
                case "Power On Off Test":
                    filename = Constants.XML_ROUTER_POWERONOFF;
                    break;
                case "Interoperability Test":
                    filename = Constants.XML_ROUTER_INTEROPERABILITY;
                    break;
                case Constants.TESTITEM_ROUTER_ChamberPerformance:
                    filename = Constants.XML_ROUTER_ChamberPerformance;
                    break;
                default:
                    return;
            }

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = filename;
            openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\config\";
            // Set filter for file extension and default file extension
            openFileDialog1.Filter = "XML file|*.xml";

            // If the file name is not an empty string open it for opening.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName != "")
            {
                switch (sTestItem)
                {
                    case "RvR Test":
                        ReadXmlRvRTest(openFileDialog1.FileName);
                        break;
                    case "RvR Turn Test":
                        ReadXmlRvRTurnTest(openFileDialog1.FileName);
                        break;
                    case "Power On Off Test":
                        ReadXmlPowerOnOff(openFileDialog1.FileName);
                        break;
                    case "Interoperability Test":
                        //ReadXmlInteroperability(openFileDialog1.FileName);
                        break;
                    case Constants.TESTITEM_ROUTER_ChamberPerformance:
                        readXmlRouterChamberPerformanceTest(openFileDialog1.FileName);
                        break;
                    default:
                        return;

                }
            }
        }

        /// <summary>
        /// This private function call is released all resources when called
        /// </summary>
        private void CloseAllResources()
        {
            /* Stop thread */
            if (threadRouterFT != null)
                threadRouterFT.Abort();

            if (threadRvRFunctionTest != null)
                threadRvRFunctionTest.Abort();

            if (threadRouterChamberPerformanceTestFT != null)
                threadRouterChamberPerformanceTestFT.Abort();

            if (cs_BrowserChamberPerformanceTest != null)
                cs_BrowserChamberPerformanceTest.Close_WebDriver();


            GC.Collect();

        }
       
        private void interOperabilityToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

        }


        public class Constants
        {
            /* Define XML file name */
            public const string XML_ROUTER_RvR              = "Router_RvR.xml";
            public const string XML_ROUTER_POWERONOFF       = "Router_PowerOnOff.xml";
            public const string XML_ROUTER_RvRTURN          = "Router_RvRTurn.xml";
            public const string XML_ROUTER_INTEROPERABILITY = "Router_Interoperability.xml";
            public const string XML_ROUTER_THROUGHPUT       = "Throuoghput.xml";
            public const string XML_ROUTER_WebGUI           = "WebGUI.xml";
            public const string XML_ROUTER_USBStorage       = "USBStorage.xml";
            public const string XML_ROUTER_GUI = "GUI.xml";
            public const string XML_ROUTER_ChamberPerformance = "Router_ChamberPerformance.xml";
            public const string XML_ROUTER_Integration = "Router_Integration.xml";
            
            /* Define Test Item Name */
            public const string TESTITEM_ROUTER_RvR = "RvR_Test";
            public const string TESTITEM_ROUTER_POWERONOFF = "Power On Off Test";
            public const string TESTITEM_ROUTER_RvRTURN = "RvR_TurnTest";
            public const string TESTITEM_ROUTER_INTEROPERABILITY = "Interoperability Test";
            public const string TESTITEM_ROUTER_THROUGHPUT = "Throughput";
            public const string TESTITEM_ROUTER_WebGUI = "WebGUI";
            public const string TESTITEM_ROUTER_USBStorage = "UsbStorage";
            public const string TESTITEM_ROUTER_ChamberPerformance = "ChamberPerformance";
            public const string TESTITEM_ROUTER_Integration = "Integration";
            


        }

        public class TestItemConstants
        {
            /* Define Test Item name */
            public const string TESTITEM_THROUGHPUT         = "Throughput Test";
            public const string TESTITEM_POWER_ONOFF        = "Power On Off Test";          
        }

        System.IO.StreamWriter WriteDebugMsg;
        System.IO.StreamWriter WriteConsoleLog;

    }
}
