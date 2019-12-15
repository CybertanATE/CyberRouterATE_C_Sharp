///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RouterChamberPerformanceFunctionTest.cs
///  Update         : 2018-10-23
///  Version        : 1.0.181023
///  Description    : 
///  Modified       : 2018-10-23 Initial version
///  History        : 2018-10-23 Initial version  
///                  
///---------------------------------------------------------------------------------------
///

using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
//using System.Windows.Forms.DataVisualization.Charting;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Xml;
using ComportClass;
using System.Data;
//using NS_CbtSnmpClass;
using System.Net;
using System.Net.Sockets;
using System.Net.NetworkInformation;
//using SnmpSharpNet;
//using NS_CbtMibAccess;
//using NS_CbtTlvFilebldrControl;
using System.IO.Ports;
using Ns_CbtFtpClient;
//using NS_CiscoInstruments;
using NS_CbtSeleniumApi;


/* User define namespace */
//using RSInstruments;
//using AgilentInstruments;
//using UART;
/* End of User define namespace */

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        /*============================================================================================*/
        /*==================== Global Paramter and Delegate Function  Declaration ====================*/
        /*============================================================================================*/
        #region

        //===Excel===
        Excel.Application xls_excelAppRouterChamberPerformanceTest;
        Excel.Workbook xls_excelWorkBookRouterChamberPerformanceTest;
        Excel.Worksheet xls_excelWorkSheetRouterChamberPerformanceTest;
        Excel.Range xls_excelRangeRouterChamberPerformanceTest;

        //int i_BandExcelReportColumnChamberPerformanceTest = 4;
        //int i_ModeExcelReportColumnChamberPerformanceTest = 5;
        //int i_ChannelExcelReportColumnChamberPerformanceTest = 6;
        //int i_BandWidthExcelReportColumnChamberPerformanceTest = 7;
        //int i_SecurityExcelReportColumnChamberPerformanceTest = 8;
        //int i_SecurityKeyExcelReportColumnChamberPerformanceTest = 9;

        int i_BandColumnChamberPerformanceTest = 5;
        int i_ModeColumnChamberPerformanceTest = 6;
        int i_ChannelColumnChamberPerformanceTest = 7;
        int i_BandWidthColumnChamberPerformanceTest = 8;
        int i_SecurityColumnChamberPerformanceTest = 9;
        int i_SecurityKeyColumnChamberPerformanceTest = 10;
        int i_TxResultColumnChamberPerformanceTest = 11;
        int i_RxResultColumnChamberPerformanceTest = 12;
        int i_BiResultColumnChamberPerformanceTest = 13;
        int i_CommentColumnChamberPerformanceTest = 14;
        int i_LogColumnChamberPerformanceTest = 15;


        //===ComPort===        
        Comport comPortRouterChamberPerformanceTest;
        Comport2 comPortCableSwIntegrationSwtich;

        //===Thread===
        Thread threadRouterChamberPerformanceTestFT;
        bool bRouterChamberPerformanceTestFTThreadRunning = false;

        //=============
        int m_PositionRouterChamberPerformanceTest = 16;
        int m_LoopRouterChamberPerformanceTest = 1;
        string m_RouterChamberPerformanceTestMainFolder = "";
        string m_RouterChamberPerformanceTestSubFolder = "";

        RouterDutsSetting[] st_Duts;
        string[] sa_GuiScriptCommandSequenceChamberPerformanceTest;
        string[] sa_GuiScriptValueSequenceChamberPerformanceTest;
        
        //===Selenium===
        CBT_SeleniumApi cs_BrowserChamberPerformanceTest = null;
        string s_CurrentUrlChamberPerformanceTest = string.Empty;
        string s_CurrentGuiAccessContent = string.Empty;
        

        //====Email===        
        string s_EmailSenderID = string.Empty;
        string s_EmailSenderPassword = string.Empty;
        string[] s_EmailReceivers = null;

        //===Delegate===
        /* Declare delegate prototype */
        public delegate void showRouterChamberPerformanceTestGUIDelegate();

        #endregion

        /*============================================================================================*/
        /*========================== Controller Event Function Area   ================================*/
        /*============================================================================================*/
        #region

        private void chamberPerformanceTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sTestItem = Constants.TESTITEM_ROUTER_ChamberPerformance;

            ToggleToolStripMenuItem(false);
            tabControl_RouterStartPage.Hide();

            chamberPerformanceTestToolStripMenuItem.Checked = true;

            //            //Hide tpCableIntegrationTemp TabPage
            //            tpCableIntegrationTemp.Parent = null;

            tabControl_ChamberPerformance.Show();

            /* Preload settings */
            string xmlFile = System.Windows.Forms.Application.StartupPath + "\\config\\" + Constants.XML_ROUTER_ChamberPerformance;
            Debug.WriteLine(File.Exists(xmlFile) ? "File exist" : "File is not existing");

            if (!File.Exists(xmlFile))
            {
                writeXmlDefaultRouterChamberPerformanceTest(xmlFile);
            }

            readXmlRouterChamberPerformanceTest(xmlFile);
            tsslMessage.Text = tabControl_ChamberPerformance.TabPages[tabControl_ChamberPerformance.SelectedIndex].Text + " Control Panel";

            /* Initial Test Condition */
            InitRouterChamberPerformanceDutsSetting();
            /* Load Duts Default Setting */
            //LoadRouterDutsSettingDataGridViewChamberPerformance();

            string filename = System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterChamberPerformanceDutsSetting.xml";

            if (File.Exists(filename))
                readXmlRouterChamberPerformanceDutsSetting(filename);
            //InitCableSwIntegrationVeriwaveTestCondition();
            //LoadCableSwIntegrationDataGridView();
        }

        private void tabControl_ChamberPerformance_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPage.Name == "tpRouterChamberPerformanceDutsSetting")
            {// Get Current comport when Data Sets page is selected
                if (cboxRouterChamberPerformanceDutsSettingComPort.Items.Count == 0)
                {
                    string[] Ports = SerialPort.GetPortNames();
                    cboxRouterChamberPerformanceDutsSettingComPort.SelectedIndex = -1;
                    foreach (string port in Ports)
                    {
                        cboxRouterChamberPerformanceDutsSettingComPort.Items.Add(port);
                        // set index value to 0 if serial device finding
                        cboxRouterChamberPerformanceDutsSettingComPort.SelectedIndex = 0;
                    } //End of foreach
                } //End of if (cboxRouterChamberPerformanceDutsSettingComPort.Items.Count == 0)
            } //End of if (e.TabPage.Name == "tpRouterChamberPerformanceDutsSetting")
        }

        private void tabControl_ChamberPerformance_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (hasDeleteButton)
            {
                if (btnRouterChamberPerformanceDutsSettingEditSetting.Text == "Cancel")
                {
                    btnRouterChamberPerformanceDutsSettingEditSetting.Text = "Edit";
                    hasDeleteButton = false;
                    dgvRouterChamberPerformanceDutsSettingData.Columns.Remove("Action");
                }
            }
        }

        private void btnRouterChamberPerformanceConfigurationLineMessageXmlFileName_Click(object sender, EventArgs e)
        {
            string filename = @"LineGroup.xml";
            string sFilter = "XML file|*.xml|All files (*.*)|*.*";
            string sInitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";

            txtRouterChamberPerformanceConfigurationLineMessageXmlFileName.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }

        private void btnRouterChamberPerformanceConfigurationTestConditionExcelFile_Click(object sender, EventArgs e)
        {
            string filename = @"RouterChamberPerformanceTestCondition.xlsx";
            string sFilter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            string sInitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";

            txtRouterChamberPerformanceConfigurationTestConditionExcelFileName.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }

        private void btnRouterChamberPerformanceConfigurationChariotFolder_Click(object sender, EventArgs e)
        {
            string filename = @"IxChariot.exe";
            string sFilter = "Exe file (*.exe)|*.exe|All files (*.*)|*.*";
            string sInitialDirectory = @"C:\Program Files\Ixia\IxChariot\";

            txtRouterChamberPerformanceConfigurationChariotFolder.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }

        private void btnRouterChamberPerformanceFunctionTestSaveLog_Click(object sender, EventArgs e)
        {
            SaveLog(txtRouterChamberPerformanceFunctionTestInformation);
        }

        private void btnRouterChamberPerformanceFunctionTestRun_Click(object sender, EventArgs e)
        {
            //            modelInfo.ModelName = txtRouterChamberPerformanceFunctionTestName.Text;
            //            modelInfo.SN = txtRouterChamberPerformanceFunctionTestSerialNumber.Text;
            //            modelInfo.SwVersion = txtRouterChamberPerformanceFunctionTestSwVersion.Text;
            //            modelInfo.HwVersion = txtRouterChamberPerformanceFunctionTestHwVersion.Text;

            /* Prevent double-click from double firing the same action */
            btnRouterChamberPerformanceFunctionTestRun.Enabled = false;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            if (bRouterChamberPerformanceTestFTThreadRunning == false)
            {
                //                //if (dgvRouterChamberPerformanceDutsSettingData.RowCount > 1)
                //                //{
                //                ////    if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                //                ////        return;
                //                ////    else
                //                ////    {
                //                ////        DataTable dt = (DataTable)dgvRouterChamberPerformanceDutsSettingData.DataSource;
                //                //    dgvRouterChamberPerformanceDutsSettingData.Rows.Clear();
                //                ////        dgvRouterChamberPerformanceDutsSettingData.DataSource = dt;
                //                ////    }
                //                //}            

                //                //readXmlRouterChamberPerformanceDutsSetting(System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterChamberPerformanceDutsSetting.xml");

                txtRouterChamberPerformanceFunctionTestInformation.Text = "";               

                if (!ParameterCheckRouterChamberPerformanceTest())
                {
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
                    return;
                }

                if (!InitialRouterChamberPerformanceTest())
                {
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
                    return;
                }

                ///* Create COM port */
                //ReadComportSettingAndInitial();
                //comPortRouterChamberPerformanceTest = new Comport();

                /* Initial Selenium */
                //cs_BrowserChamberPerformanceTest = new CBT_SeleniumApi();

                /* Create sub folder for saving the report data */
                m_RouterChamberPerformanceTestMainFolder = createReportMainFolderRouterChamberPerformance();

                bRouterChamberPerformanceTestFTThreadRunning = true;
                /* Disable all controller */
                ToggleRouterChamberPerformanceFunctionTestController(false);

                btnRouterChamberPerformanceFunctionTestRun.Text = "Stop";
                
                threadRouterChamberPerformanceTestFT =new Thread(new ThreadStart(DoRouterChamberPerformanceFunctionTest));
                threadRouterChamberPerformanceTestFT.Name = "";
                threadRouterChamberPerformanceTestFT.Start();
            }
            else
            {
                tsslMessage.Text = "Function Test Control Panel";
                bRouterChamberPerformanceTestFTThreadRunning = false;
            }
                        

            /* Button release */
            Thread.Sleep(3000);
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
        }

        //        private void btnCableSwIntegrationConfigurationReportFolder_Click(object sender, EventArgs e)
        //        {
        //            FolderBrowserDialog folder = new FolderBrowserDialog();
        //            folder.ShowDialog();
        //            txtCableSwIntegrationConfigurationReportFolder.Text = folder.SelectedPath;
        //        }

        //        //private void btnRouterChamberPerformanceFunctionTestMibExcelFile_Click(object sender, EventArgs e)
        //        //{
        //        //    string filename = @"MibReadWriteConfigFileTestCondition.xlsx";
        //        //    // Displays a SaveFileDialog so the user can select the exported test condition excel file
        //        //    OpenFileDialog openFileDialog1 = new OpenFileDialog();
        //        //    //openFileDialog1.Multiselect = true;
        //        //    openFileDialog1.FileName = filename;
        //        //    openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\importData\";
        //        //    // Set filter for file extension and default file extension
        //        //    openFileDialog1.Filter = "XLSX file|*.xlsx";

        //        //    // If the file name is not an empty string open it for opening.
        //        //    if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName != "")
        //        //    {
        //        //        //string files = string.Empty;

        //        //        //foreach (string s in openFileDialog1.FileNames)
        //        //        //{
        //        //        //    files = files + s;
        //        //        //    files += ";";
        //        //        //}

        //        //        //txtRouterChamberPerformanceFunctionTestMibExcelFile.Text = files;

        //        //        txtRouterChamberPerformanceFunctionTestMibExcelFile.Text = openFileDialog1.FileName;
        //        //    }
        //        //}


        

        
        
        
        
        

        //                if (comPortRouterChamberPerformanceTest.isOpen() != true)
        //                {
        //                    MessageBox.Show("COM port is not ready!");
        //                    System.Windows.Forms.Cursor.Current = Cursors.Default;
        //                    btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
        //                    return;
        //                }

        //                if (chkCableSwIntegrationConfigurationUseSwitch.Checked)
        //                {
        //                    ///* Create COM port */
        //                    ReadComport2SettingAndInitial();

        //                    ///* Create COM port 2 for switch control */
        //                    comPortCableSwIntegrationSwtich = new Comport2();

        //                    if (comPortCableSwIntegrationSwtich.isOpen() != true)
        //                    {
        //                        MessageBox.Show("COM port 2 is not ready!");
        //                        System.Windows.Forms.Cursor.Current = Cursors.Default;
        //                        btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
        //                        return;
        //                    }                   
        //                }

        
       


        

        //                if (rbtnRouterChamberPerformanceFunctionTestRightNow.Checked)
        //                {
        //                    ForTestRouterChamberPerformanceTest();
        //                }

        //                string sDateNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        //                string str = "Now is: " + sDateNow;
        //                SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

        //                str = "Setting Time is: " + i_SetCheckHourTime + ":" + i_SetCheckMinuteTime + ":" + i_SetCheckSecondTime;
        //                SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

        //                str = "Waiting for time up....";
        //                SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

        //                bRouterChamberPerformanceTestFTThreadRunning = true;
        //                /* Disable all controller */
        //                ToggleRouterChamberPerformanceFunctionTestController(false);

        //                btnRouterChamberPerformanceFunctionTestRun.Text = "Stop";


        //                t_Timer = new System.Timers.Timer();
        //                t_Timer.Enabled = true;
        //                t_Timer.Interval = 1000;
        //                t_Timer.Start();
        //                t_Timer.Elapsed += new System.Timers.ElapsedEventHandler(RouterChamberPerformanceTestTimer1_Elapsed); 

        //                /* Create sub folder for saving the report data */
        //                //m_RouterChamberPerformanceTestSubFolder = createRouterChamberPerformanceTestSubFolder(txtRouterChamberPerformanceFunctionTestName.Text);                       

        //                //
        //                //threadRouterChamberPerformanceTestFT = new Thread(new ThreadStart(DoRouterChamberPerformanceFunctionTest));
        //                //threadRouterChamberPerformanceTestFT.Name = "";
        //                //threadRouterChamberPerformanceTestFT.Start();
        //            }
        
        //        }

        #endregion

        /*============================================================================================*/
        /*=================================== Main Function Area   ===================================*/
        /*============================================================================================*/
        #region
        private void DoRouterChamberPerformanceFunctionTest()
        {
            int m_totalTimes = 1;   // Default test time
            m_LoopRouterChamberPerformanceTest = 1; // Reset loop counter
            if (chkRouterChamberPerformanceFunctionTestScheduleOnOff.Checked)
                m_totalTimes = Convert.ToInt32(nudRouterChamberPerformanceFunctionTestTimes.Value);

            /* Start Test Loop */
            do
            {
                if (bRouterChamberPerformanceTestFTThreadRunning == false)
                {
                    bRouterChamberPerformanceTestFTThreadRunning = false;
                    MessageBox.Show("Abort test", "Error");
                    this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
                    threadRouterChamberPerformanceTestFT.Abort();
                    // Never go here ...
                }               

                RouterChamberPerformanceTestMainFunction();

                m_totalTimes--;
                m_LoopRouterChamberPerformanceTest++;
            } while (m_totalTimes != 0);

            bRouterChamberPerformanceTestFTThreadRunning = false;
            this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));      
        }

        private bool RouterChamberPerformanceTestMainFunction()
        {
            string str = string.Empty;
            string sExcelFile = string.Empty;
            string sInputFile = string.Empty;
            string sOutputFile = string.Empty;
            string sComment = string.Empty;

            string[] sReport = new string[3];

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            /* Initial Value */
            m_PositionRouterChamberPerformanceTest = 16;
            //testConfigCableSwIntegrationTest = null;

            str = String.Format("======= Run Main Function ======");
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            /* Read Duts Setting to run Main function */
            for (int DutsRow = 0; DutsRow < st_DutsSetting.Length; DutsRow++)
            {
                if (bRouterChamberPerformanceTestFTThreadRunning == false)
                {
                    bRouterChamberPerformanceTestFTThreadRunning = false;
                    MessageBox.Show("Abort test", "Error");
                    this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
                    threadRouterChamberPerformanceTestFT.Abort();
                    // Never go here ...
                }    

                sInputFile = string.Empty;
                sOutputFile = string.Empty;
                sComment = string.Empty;

                m_PositionRouterChamberPerformanceTest = 16;

                modelInfo.ModelName = st_DutsSetting[DutsRow].ModelName;
                modelInfo.SN = st_DutsSetting[DutsRow].SerialNumber;
                modelInfo.SwVersion = st_DutsSetting[DutsRow].SwVersion;
                modelInfo.HwVersion = st_DutsSetting[DutsRow].HwVersion;
                
                string sChariotTxTstFile = sa_RounterDutsSetting[DutsRow, 13];
                string sChariotRxTstFile = sa_RounterDutsSetting[DutsRow, 14];
                string sChariotBiTstFile = sa_RounterDutsSetting[DutsRow, 15];
                string sRouterIpAddress = st_DutsSetting[DutsRow].IpAddress;
                
                string sGuiScriptExcelFile = st_DutsSetting[DutsRow].GuiScriptExcelFile;
                string sExcelFileName = "ChamberPerformanceTest";
                //string subFolder;
                string sLogFile;

                str = String.Format("Run DUT: {0}", modelInfo.ModelName);
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                /* Create Folder */
                str = "Create Folder";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                try
                {
                    m_RouterChamberPerformanceTestSubFolder = createSubFolderRouterChamberPerformance(m_RouterChamberPerformanceTestMainFolder);
                }
                catch (Exception ex)
                {
                    str = "Create Folder Failed. Exception: " + ex.ToString();
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    str = "Create Folder Failed.";
                    sLogFile = m_RouterChamberPerformanceTestMainFolder + "\\" + modelInfo.ModelName + "_HeaderLog_" + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".txt";
                    WriteFailMsgToExcelAndSendEamilChamberPerformanceTest(str, sLogFile);
                    SendLineAlarmChamberPerformanceTest(str);                    
                    continue;
                }
                
                str = "Finished!";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                sLogFile = m_RouterChamberPerformanceTestSubFolder + "\\" + "HeaderLog_" + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".txt";

                #region  marked area
                ///* Read Test Condition xls File */
                //if (!File.Exists(txtRouterChamberPerformanceConfigurationTestConditionExcelFileName.Text))
                //{


                //    m_PositionRouterChamberPerformanceTest = 16;
                //    continue;
                //}


                ///* Read Steps Condition xls File */
                #endregion marked area

                /* Create Folder */
                str = "Read Gui Script";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                /* Check Gui Script File Existence */
                if (!File.Exists(sGuiScriptExcelFile))
                { //Fail. Send Log and Line Alarm
                    str = "Gui Script File doesn't Exist!!";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                                        
                    WriteFailMsgToExcelAndSendEamilChamberPerformanceTest(str, sLogFile);
                    SendLineAlarmChamberPerformanceTest(str);

                    //m_PositionRouterChamberPerformanceTest = 16;
                    continue;
                }

                /* Read Gui Script */
                if (!ConvertExcelToDualArray(sGuiScriptExcelFile, ref sa_RounterGuiScript))
                {//Fail. Send Log and Line Alarm
                    str ="Read Gui Script File Failed!!";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    WriteFailMsgToExcelAndSendEamilChamberPerformanceTest(str, sLogFile);
                    SendLineAlarmChamberPerformanceTest(str);

                    //m_PositionRouterChamberPerformanceTest = 16;
                    continue;                  
                }
                
                str = "Finished!";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                
                string savePath = createReportExcelFileChamberPerformanceTest(m_RouterChamberPerformanceTestSubFolder, modelInfo.ModelName, m_LoopRouterChamberPerformanceTest, sExcelFileName);
                //string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
                //string savePath = createCableSwIntegrationTestSaveExcelFile(m_CableSwIntegrationTestSubFolder, cableSwIntegrationDutDataSets[i].SkuModel, m_LoopCableSwIntegrationTest, sExcelFileName);
                /* initial Excel component */
                initialExcelRouterChamberPerformanceTest(savePath);

                try
                {
                    /* Fill Loop, File Name in Excel */
                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[9, 3] = modelInfo.SN;
                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 3] = modelInfo.SwVersion;
                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[11, 3] = modelInfo.HwVersion;
                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[4, 9] = m_LoopRouterChamberPerformanceTest;
                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[5, 9] = Path.GetFileName(sGuiScriptExcelFile);
                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[6, 9] = Path.GetFileName(sChariotTxTstFile);
                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 9] = Path.GetFileName(sChariotRxTstFile);
                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[8, 9] = Path.GetFileName(sChariotBiTstFile);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }

                /* Run Main Condition for Excel Data */
                for (int rowCondition = 0; rowCondition < sa_RounterTestCondition.GetLength(0); rowCondition++)
                {
                    if (bRouterChamberPerformanceTestFTThreadRunning == false)
                    {
                        bRouterChamberPerformanceTestFTThreadRunning = false;
                        MessageBox.Show("Abort test", "Error");
                        this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
                        threadRouterChamberPerformanceTestFT.Abort();
                        // Never go here ...
                    }

                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

                    if (sa_RounterTestCondition[rowCondition, 0] == "" || sa_RounterTestCondition[rowCondition, 0] == null)
                    {
                        continue;
                    }

                    if (sa_RounterTestCondition[rowCondition, 0].ToLower().Trim() != "v")
                    {
                        continue;
                    }

                    str = String.Format("======= Start Condition Test ======");
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    str = "Model Name : " + modelInfo.ModelName;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    str = "Run Test Item " + rowCondition.ToString();
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    str = "Index is " + sa_RounterTestCondition[rowCondition, 1];
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });


                    str = "Function Name is " + sa_RounterTestCondition[rowCondition, 2];
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    sLogFile = "Log_" + sa_RounterTestCondition[rowCondition, 1] + "_" + sa_RounterTestCondition[rowCondition, 2] + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";

                    try
                    {
                        sLogFile = MakeFilenameValid(sLogFile);
                    }
                    catch (Exception)
                    {
                        sLogFile = "Log_" + sa_RounterTestCondition[rowCondition, 1] + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
                    }

                    sLogFile = m_RouterChamberPerformanceTestSubFolder + "\\" + sLogFile;
                    sReport[2] = sLogFile;
                    sReport[0] = "";
                    sReport[1] = "";

                    st_WifiParameter = new WifiParameter();
                    st_WifiParameter.Band = sa_RounterTestCondition[rowCondition, i_BandColumnChamberPerformanceTest-1];
                    st_WifiParameter.Mode = sa_RounterTestCondition[rowCondition, i_ModeColumnChamberPerformanceTest - 1];
                    st_WifiParameter.Channel = sa_RounterTestCondition[rowCondition, i_ChannelColumnChamberPerformanceTest - 1];
                    st_WifiParameter.BandWidth = sa_RounterTestCondition[rowCondition, i_BandWidthColumnChamberPerformanceTest - 1];

                    if (sa_RounterTestCondition[rowCondition, i_SecurityColumnChamberPerformanceTest - 1] != null && sa_RounterTestCondition[rowCondition, i_SecurityColumnChamberPerformanceTest - 1] != "")
                    {
                        string[] s = sa_RounterTestCondition[rowCondition, i_SecurityColumnChamberPerformanceTest - 1].Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);

                        st_WifiParameter.Security = s[0];
                        if (s.Length > 1)
                            st_WifiParameter.SecurityMode = s[1];
                    }
                    //st_WifiParameter.SecurityIndex = sa_RounterTestCondition[rowCondition, 1];
                    st_WifiParameter.SecurityKey = sa_RounterTestCondition[rowCondition, i_SecurityKeyColumnChamberPerformanceTest-1];

                    try
                    {
                        /* Fill Loop, PowerLevel and Constellation in Excel */
                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = sa_RounterTestCondition[rowCondition, 1];
                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3] = sa_RounterTestCondition[rowCondition, 2];
                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 4] = sa_RounterTestCondition[rowCondition, 3];
                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5] = sa_RounterTestCondition[rowCondition, 4];
                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 6] = sa_RounterTestCondition[rowCondition, 5];
                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7] = sa_RounterTestCondition[rowCondition, 6];
                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 8] = sa_RounterTestCondition[rowCondition, 7];
                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 9] = sa_RounterTestCondition[rowCondition, 8];
                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 10] = sa_RounterTestCondition[rowCondition, 9];
                        //xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 11] = sa_RounterTestCondition[rowCondition, 4];
                        //xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 12] = sa_RounterTestCondition[rowCondition, 4];
                        //xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 13] = sa_RounterTestCondition[rowCondition, 4];
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex);
                    }                    

                    //string sOutputPath = System.Windows.Forms.Application.StartupPath + @"\report\" + Path.GetFileName(m_RouterChamberPerformanceTestSubFolder);
                    string sResultValue = string.Empty;
                    string sOutputHead = @"\Model_Band_Mode_Chanchannel_Security_secType_";
                    /* Replace the _outpurFile with WifiBasic and Attenuation Value */
                    //sOutputHead = sOutputHead.Replace("outpurPATH", sOutputPath);
                    sOutputHead = sOutputHead.Replace("Model", modelInfo.ModelName);
                    sOutputHead = sOutputHead.Replace("Band", st_WifiParameter.Band);
                    sOutputHead = sOutputHead.Replace("Mode", st_WifiParameter.Mode);
                    sOutputHead = sOutputHead.Replace("channel", st_WifiParameter.Channel);
                    sOutputHead = sOutputHead.Replace("Security", st_WifiParameter.Security);
                    sOutputHead = sOutputHead.Replace("secType", st_WifiParameter.SecurityMode);

                    sOutputHead = MakeFilenameValid(sOutputHead);
                    string t = sOutputHead;
                    sOutputFile = m_RouterChamberPerformanceTestSubFolder +"\\" + sOutputHead + "Tx_" + Path.GetFileName(sInputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");

                    //sOutputFile = m_RouterChamberPerformanceTestMainFolder + sOutputHead + "Tx_" + Path.GetFileName(sInputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");


                    ////Check Wifi Parameter????
                    if(st_WifiParameter.Band == null || st_WifiParameter.Band == "" ||
                        st_WifiParameter.Mode == null || st_WifiParameter.Mode == "" ||
                        st_WifiParameter.Channel == null || st_WifiParameter.Channel == "" ||
                        st_WifiParameter.BandWidth == null || st_WifiParameter.BandWidth == "" ||
                        st_WifiParameter.Security == null || st_WifiParameter.Security == "" ||
                        st_WifiParameter.Band == null || st_WifiParameter.Band == "")
                    {                        
                        sReport[0] = "Error";
                        sReport[1] = "Wifi Parameter couldn't be Empty!";
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });

                        RouterChamberPerformanceTestReportData(i_TxResultColumnChamberPerformanceTest, sReport);

                        File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
                        //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

                        m_PositionRouterChamberPerformanceTest++;
                        continue;
                            
                    }

                    if(st_WifiParameter.Security.ToLower() != "none")
                    {
                        //if(st_WifiParameter.SecurityMode == null || st_WifiParameter.SecurityMode == "" ||
                        if(st_WifiParameter.SecurityKey == null || st_WifiParameter.SecurityKey == "")
                        {
                            sReport[0] = "Error";
                            sReport[1] = "Wifi Security Parameter couldn't be Empty!";
                            Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                            Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                            RouterChamberPerformanceTestReportData(i_TxResultColumnChamberPerformanceTest, sReport);

                            File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
                            //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

                            m_PositionRouterChamberPerformanceTest++;
                            continue;
                        }
                    }

                    /*Initial Setting Sequence */

                    sa_GuiScriptCommandSequenceChamberPerformanceTest[8] = "";
                    sa_GuiScriptCommandSequenceChamberPerformanceTest[9] = "";
                    sa_GuiScriptCommandSequenceChamberPerformanceTest[10] = "";
                    sa_GuiScriptCommandSequenceChamberPerformanceTest[11] = "";
                    sa_GuiScriptCommandSequenceChamberPerformanceTest[12] = "";                    

                    switch (st_WifiParameter.Security.ToLower())
                    {
                        case "none":
                            {
                                sa_GuiScriptCommandSequenceChamberPerformanceTest[7] = "Setting Wifi Security None 5G";
                                sa_GuiScriptValueSequenceChamberPerformanceTest[7] = st_WifiParameter.Security;
                            }
                            break;
                        case "wpa2 personal":
                            {
                                sa_GuiScriptCommandSequenceChamberPerformanceTest[7] = "Setting Wifi Security PSK 5G";
                            }
                            break;
                        case "wpa2 enterprise":
                            {
                                sa_GuiScriptCommandSequenceChamberPerformanceTest[7] = "Setting Wifi Security ENT KEY 5G";
                                sa_GuiScriptCommandSequenceChamberPerformanceTest[8] = "Setting Wifi Security Radius IP.1 5G";
                                sa_GuiScriptCommandSequenceChamberPerformanceTest[9] = "Setting Wifi Security Radius IP.2 5G";
                                sa_GuiScriptCommandSequenceChamberPerformanceTest[10] = "Setting Wifi Security Radius IP.3 5G";
                                sa_GuiScriptCommandSequenceChamberPerformanceTest[11] = "Setting Wifi Security Radius IP.4 5G";
                                sa_GuiScriptCommandSequenceChamberPerformanceTest[12] = "Setting Wifi Security Radius Port 5G";                         
                            }
                            break;
                        default:
                            {                                
                                sReport[0] = "Error";
                                sReport[1] = "Unsupport Security Mode!";
                                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                                RouterChamberPerformanceTestReportData(i_TxResultColumnChamberPerformanceTest, sReport);

                                File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
                                //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

                                m_PositionRouterChamberPerformanceTest++;
                                continue;
                            }
                            break;
                    }

                    if(st_WifiParameter.Band.ToLower().IndexOf("2.4") >=0)
                    {//2.4G
                        
                        /* Set SSID */
                        st_WifiParameter.SSID = sa_RounterDutsSetting[DutsRow, 6];

                        for (int i = 0; i < sa_GuiScriptCommandSequenceChamberPerformanceTest.Length; i++)
                        {
                            sa_GuiScriptCommandSequenceChamberPerformanceTest[i] = sa_GuiScriptCommandSequenceChamberPerformanceTest[i].Replace("5G", "2.4G");
                        }                        
                    }
                    else
                    {//5G
                        /* Set SSID */
                        st_WifiParameter.SSID = sa_RounterDutsSetting[DutsRow, 7];

                        for (int i = 0; i < sa_GuiScriptCommandSequenceChamberPerformanceTest.Length; i++)
                        {
                            sa_GuiScriptCommandSequenceChamberPerformanceTest[i] = sa_GuiScriptCommandSequenceChamberPerformanceTest[i].Replace("2.4G", "5G");
                        }                                               
                    }
                    
                    sa_GuiScriptValueSequenceChamberPerformanceTest[1] = st_WifiParameter.SSID;                    
                    sa_GuiScriptValueSequenceChamberPerformanceTest[2] = st_WifiParameter.Mode;                    
                    sa_GuiScriptValueSequenceChamberPerformanceTest[3] = st_WifiParameter.BandWidth;
                    sa_GuiScriptValueSequenceChamberPerformanceTest[4] = st_WifiParameter.Channel;
                    sa_GuiScriptValueSequenceChamberPerformanceTest[5] = "None";
                    sa_GuiScriptValueSequenceChamberPerformanceTest[6] = st_WifiParameter.Security;
                    sa_GuiScriptValueSequenceChamberPerformanceTest[7] = st_WifiParameter.SecurityKey;
                    if(st_WifiParameter.Security.ToLower() == "none")
                    {
                        sa_GuiScriptValueSequenceChamberPerformanceTest[6] = st_WifiParameter.Security;
                        st_WifiParameter.SecurityKey = st_WifiParameter.Security;
                        st_WifiParameter.SecurityMode = st_WifiParameter.Security;
                    }                    
                                        
                    

                    str = "";
                    if (!SettingRouterGuiRouterChamberPerformanceTest(sRouterIpAddress, st_WifiParameter, sa_RounterGuiScript))
                    {
                        str = "Setting GUI Script Fail!! ";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                        sReport[0] = "Fail";
                        sReport[1] = str;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });

                        RouterChamberPerformanceTestReportData(i_TxResultColumnChamberPerformanceTest, sReport);

                        File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
                        //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

                        m_PositionRouterChamberPerformanceTest++;
                        continue;
                    } 
   
                    ///* For test */
                    //st_WifiParameter.SSID = "Miyophone";
                    //st_WifiParameter.Security = "WPA2PSK";
                    //st_WifiParameter.SecurityMode = "AES";
                    //st_WifiParameter.SecurityKey = "qwertyuio";

                    //st_WifiParameter.SSID = "Linksys29719";
                    //st_WifiParameter.Security = "None";
                    //st_WifiParameter.SecurityMode = "AES";
                    //st_WifiParameter.SecurityKey = "crmew0oh3w";

                    //st_WifiParameter.SSID = "Linksys29719";
                    //st_WifiParameter.Security = "WEP";
                    //st_WifiParameter.SecurityMode = "1";
                    //st_WifiParameter.SecurityKey = "1234567890";

                    //st_WifiParameter.SSID = "Miyohtc";
                    //st_WifiParameter.Security = "WPA2PSK";
                    //st_WifiParameter.SecurityMode = "AES";
                    //st_WifiParameter.SecurityKey = "qwertyuio";

                    //Control Remote side to connect to Dut
                    str = "Connect to Remote Server... ";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    if (!RemoteConnectionRouterChamberPerformanceTest(st_WifiParameter, "WifiConnection", ref sReport))
                    {
                        str = "Remote Connection Fail!! ";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                        sReport[0] = "Fail";
                        sReport[1] = str;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });

                        RouterChamberPerformanceTestReportData(i_TxResultColumnChamberPerformanceTest, sReport);

                        File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
                        //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

                        m_PositionRouterChamberPerformanceTest++;
                        continue;
                    }

                    str = "Succeed!";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    #region Ping Client Card
                    //Ping Client Card
                    string sIp = mtbRouterChamberPerformanceConfigurationClientCardIpAddress.Text;

                    str = "Ping Client Card, IP Address: " + sIp;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    IPAddress ipTemp;
                    if (!IPAddress.TryParse(mtbRouterChamberPerformanceConfigurationClientCardIpAddress.Text, out ipTemp))
                    {
                        str = "IP Invalid!! ";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                        sReport[0] = "Fail";
                        sReport[1] = str;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });

                        RouterChamberPerformanceTestReportData(i_TxResultColumnChamberPerformanceTest, sReport);

                        File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
                        //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

                        m_PositionRouterChamberPerformanceTest++;
                        continue;
                    }

                    if (!QuickPingChamberPerformanceTest(sIp, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
                    {
                        str = "Ping Failed!! ";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                        sReport[0] = "Fail";
                        sReport[1] = str;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });

                        RouterChamberPerformanceTestReportData(i_TxResultColumnChamberPerformanceTest, sReport);

                        File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
                        //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

                        m_PositionRouterChamberPerformanceTest++;
                        continue;                      
                    }
                    #endregion

                    #region Run Chariot

                    str = "Start to run Chariot";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    sInputFile = sChariotTxTstFile;
                    str = "Run Chariot Tx tst File: " + sInputFile;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                                       
                    
                    if (!ChariotFunctionRouterChamberPerformanceTest(sInputFile, sOutputFile, ref sResultValue, ref sComment))
                    {
                        sReport[0] = "Fail";
                        sReport[1] = sComment;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                    }
                    else
                    {
                        sReport[0] = sResultValue;
                        str = "Tx Throughput = " + sResultValue;
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                    }

                    RouterChamberPerformanceTestReportData(i_TxResultColumnChamberPerformanceTest, sReport);

                    sInputFile = sChariotRxTstFile;
                    str = "Run Chariot Rx tst File: " + sInputFile;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    sOutputFile = m_RouterChamberPerformanceTestSubFolder + "\\" + sOutputHead + "Rx_" + Path.GetFileName(sInputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");

                    if (!ChariotFunctionRouterChamberPerformanceTest(sInputFile, sOutputFile, ref sResultValue, ref sComment))
                    {
                        sReport[0] = "Fail";
                        sReport[1] = sComment;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                    }
                    else
                    {
                        sReport[0] = sResultValue;
                        str = "Rx Throughput = " + sResultValue;
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                    }

                    RouterChamberPerformanceTestReportData(i_RxResultColumnChamberPerformanceTest, sReport);

                    sInputFile = sChariotBiTstFile;
                    str = "Run Chariot Bi tst File: " + sInputFile;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    sOutputFile = m_RouterChamberPerformanceTestSubFolder + "\\" + sOutputHead + "Bi_" + Path.GetFileName(sInputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");

                    if (!ChariotFunctionRouterChamberPerformanceTest(sInputFile, sOutputFile, ref sResultValue, ref sComment))
                    {
                        sReport[0] = "Fail";
                        sReport[1] = sComment;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                    }
                    else
                    {
                        sReport[0] = sResultValue;
                        str = "Bi Throughput = " + sResultValue;
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                    }

                    RouterChamberPerformanceTestReportData(i_BiResultColumnChamberPerformanceTest, sReport);
                    //sLogFile = "Log_" + sa_RounterTestCondition[rowCondition, 1] + "_" + sa_RounterTestCondition[rowCondition, 2] + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";

                    //try
                    //{
                    //    sLogFile = MakeFilenameValid(sLogFile);
                    //}
                    //catch (Exception)
                    //{
                    //    sLogFile = "Log_" + sa_RounterTestCondition[rowCondition, 1] + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
                    //}

                    //sLogFile = m_RouterChamberPerformanceTestSubFolder + "\\" + sLogFile;
                    //sReport[2] = sLogFile;

                    //RouterChamberPerformanceTestReportData(11, sReport);
                    
                    #endregion

                    File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);

                    m_PositionRouterChamberPerformanceTest++;
                }// End of Run Final Excel Condition

                /* Save end time and close the Excel object */
                xls_excelAppRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                xls_excelWorkBookRouterChamberPerformanceTest.Save();
                /* Close excel application when finish one job(power level) */
                closeExcelRouterChamberPerformanceTest();
                Thread.Sleep(3000);

                if (chkRouterChamberPerformanceConfigurationSendReport.Checked)
                {
                    s_EmailSenderID = txtRouterChamberPerformanceConfigurationReportEmailSenderGmailAccount.Text;
                    s_EmailSenderPassword = txtRouterChamberPerformanceConfigurationReportEmailSenderGmailPassword.Text;
                    s_EmailReceivers = txtRouterChamberPerformanceConfigurationReportEmailSendTo.Text.Split(new string[]{";", ","}, StringSplitOptions.RemoveEmptyEntries);

                    try
                    {
                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Router Chamber Performance ATE Test : " + modelInfo.ModelName + " Test Report", "Test Completed!!!", savePath);
                    }
                    catch (Exception ex)
                    {
                        str = "Send Mail Failed: " + ex.ToString();
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                    }
                }

            }//End of Run Duts Condition

            return true;
        }

        /* The function is recorded report data */
        private void RouterChamberPerformanceTestReportData(int iColumn, string[] sReport)
        {
            try
            {
                /* Fill Loop, PowerLevel and Constellation in Excel */
                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, iColumn] = sReport[0];
                if (sReport[0] != null && sReport[0] != "")
                {
                    if (sReport[0].ToLower() == "fail")
                    {
                        xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, iColumn], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, iColumn]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                }               
                
                if (sReport[1] != null || sReport[1] != "")
                {
                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, i_CommentColumnChamberPerformanceTest] = sReport[1];
                }

                string sLogFile = sReport[2];
                /* Write Console Log File as HyperLink in Excel report */
                xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, i_LogColumnChamberPerformanceTest];
                xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sLogFile, Type.Missing, "ConsoleLog", "ConsoleLog");                  
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            try
            {
                /* Save Excel */                
                xls_excelWorkBookRouterChamberPerformanceTest.Save();                
            }
            catch (Exception ex)
            {
                //str = "SavWrite Data to Excel File Exception: " + ex.ToString();
                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
            }
        }

        #endregion
        
        //======== Sub-Function of Main
        #region Parameter Check and Data Initialize

        private bool ParameterCheckRouterChamberPerformanceTest()
        {
            /* Config and Check Test Condition */
            SetText("Check Paramters...", txtRouterChamberPerformanceFunctionTestInformation);

            /* Check Duts Setting Content */
            if (dgvRouterChamberPerformanceDutsSettingData.RowCount <= 1)
            {
                MessageBox.Show("Duts Setting Data Area can't be Empty!!");
                return false;
            }

            /* Check Email Parameter */
            if (chkRouterChamberPerformanceConfigurationSendReport.Checked)
            {
                if (txtRouterChamberPerformanceConfigurationReportEmailSendTo.Text == "")
                {
                    MessageBox.Show("Email Send-to list can't be Empty!!");
                    return false;
                }

                if (txtRouterChamberPerformanceConfigurationReportEmailSenderGmailAccount.Text == "")
                {
                    MessageBox.Show("Gmail Account can't be Empty!!");
                    return false;
                }

                if (txtRouterChamberPerformanceConfigurationReportEmailSenderGmailPassword.Text == "")
                {
                    MessageBox.Show("Gmail Password can't be Empty!!");
                    return false;
                }
            }

            /* Check Line Message Parameter */
            if (chkRouterChamberPerformanceConfigurationLineMessage.Checked)
            {
                if (txtRouterChamberPerformanceConfigurationLineMessageXmlFileName.Text == "")
                {
                    MessageBox.Show("Line Group list can't be Empty!!");
                    return false;
                }

                /* Check File Existence */
                if (!File.Exists(txtRouterChamberPerformanceConfigurationLineMessageXmlFileName.Text))
                {
                    MessageBox.Show("Line Group Xml File doesn't exist!!");
                    return false;
                }
            }

            /* Check Swtich Setting */            
            if (chkRouterChamberPerformanceConfigurationUseSwitch.Checked)
            {
                if (txtRouterChamberPerformanceConfigurationSwitchUsername.Text == "")
                {
                    MessageBox.Show("Switch Login Username can't be Empty!!");
                    return false;
                }

                if (txtRouterChamberPerformanceConfigurationSwitchPassword.Text == "")
                {
                    MessageBox.Show("Switch Login Password can't be Empty!!");
                    return false;
                }
            }

            /* Check Other Setting */
            if (mtbRouterChamberPerformanceConfigurationRemoteControlServerIpAddress.Text == "")
            {
                MessageBox.Show("Remote Server IP address can't be Empty!!");
                return false;
            }

            if (nudRouterChamberPerformanceConfigurationRemoteControlServerPort.Text == "")
            {
                MessageBox.Show("Remote Server Port number can't be Empty!!");
                return false;
            }

            if (txtRouterChamberPerformanceConfigurationTestConditionExcelFileName.Text == "")
            {
                MessageBox.Show("Test Condition Excel File can't be Empty!!");
                return false;
            }

            /* Check Radius Server */
            if (mtbRouterChamberPerformanceConfigurationRadiusServerIpAddress.Text == "")
            {
                MessageBox.Show("Radius Server IP address can't be Empty!!");
                return false;
            }

            if (nudRouterChamberPerformanceConfigurationRadiusServerPort.Text == "")
            {
                MessageBox.Show("Radius Server Port number can't be Empty!!");
                return false;
            }

            /* Check File Existence */
            if (!File.Exists(txtRouterChamberPerformanceConfigurationTestConditionExcelFileName.Text))
            {
                MessageBox.Show("Test Condition Excel File doesn't exist!!");
                return false;
            }            

            SetText("Finished!", txtRouterChamberPerformanceFunctionTestInformation);

            return true;
        }

        private bool InitialRouterChamberPerformanceTest()
        {
            /* Config and Check Test Condition */
            SetText("Initial Data...", txtRouterChamberPerformanceFunctionTestInformation);

            /* Initial Test Condition */
            SetText("Initial Test Condition Array...", txtRouterChamberPerformanceFunctionTestInformation);
            if (!InitTestConditionArrayRouterChamberPerformanceTest())
                return false;            

            ///* Initial Test Condition */
            //SetText("Initial Steps Condition Array...", txtRouterChamberPerformanceFunctionTestInformation);
            //if (!InitStepsConditionArrayRouterChamberPerformanceTest())
            //    return false;            

            /* Initial Test Condition */
            SetText("Initial Duts Setting Array...", txtRouterChamberPerformanceFunctionTestInformation);
            if (!InitDusSettingArrayRouterChamberPerformanceTest())
                return false;

            SetText("Finished!", txtRouterChamberPerformanceFunctionTestInformation);

            /* Initial Test Condition */
            SetText("Initial Dut GUI Script Command Sequence...", txtRouterChamberPerformanceFunctionTestInformation);
            if (!InitGuiScriptCommandSequenceRouterChamberPerformanceTest())
                return false;

            SetText("Finished!", txtRouterChamberPerformanceFunctionTestInformation);
            
            return true;
        }

        private bool InitTestConditionArrayRouterChamberPerformanceTest()
        {
            //SetText("Initial Test Condition Array...", txtRouterChamberPerformanceFunctionTestInformation);
            string sFileName = txtRouterChamberPerformanceConfigurationTestConditionExcelFileName.Text;
            //string[,] saCondition = new string[1, 1];

            if (!ConvertExcelToDualArray(sFileName, ref sa_RounterTestCondition))
            {
                MessageBox.Show("Read Test Condition Failed!!");
                return false;
            }

            return true;
        }

        private bool InitStepsConditionArrayRouterChamberPerformanceTest()
        {
            //SetText("Initial Steps Condition Array...", txtRouterChamberPerformanceFunctionTestInformation);
            string sFileName = "";

            if (!ConvertExcelToDualArray(sFileName, ref sa_RounterStepsCondition))
            {
                MessageBox.Show("Read Steps Condition Failed!!");
                return false;
            }
            return true;
        }

        private bool InitDusSettingArrayRouterChamberPerformanceTest()
        {
            //SetText("Initial Duts Setting Array...", txtRouterChamberPerformanceFunctionTestInformation);

            if (!ConvertDatagridViewDataToDualArray(dgvRouterChamberPerformanceDutsSettingData, ref sa_RounterDutsSetting))
            {
                MessageBox.Show("Read Duts Setting Failed!!");
                return false;
            }

            st_DutsSetting = new RouterDutsSetting[dgvRouterChamberPerformanceDutsSettingData.RowCount - 1];

            for (int i = 0; i < dgvRouterChamberPerformanceDutsSettingData.RowCount - 1; i++)
            {
                st_DutsSetting[i].index = Convert.ToInt32(dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[0].Value.ToString());
                st_DutsSetting[i].ModelName = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[1].Value.ToString();
                st_DutsSetting[i].SerialNumber = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[2].Value.ToString();
                st_DutsSetting[i].SwVersion = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[3].Value.ToString();
                st_DutsSetting[i].HwVersion = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[4].Value.ToString();
                st_DutsSetting[i].IpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[5].Value.ToString();
                //st_DutsSetting[i].SSID = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[6].Value.ToString();
                st_DutsSetting[i].PcIpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[8].Value.ToString();
                st_DutsSetting[i].SwitchPort = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[9].Value.ToString();
                st_DutsSetting[i].ComportNum = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[10].Value.ToString();
                st_DutsSetting[i].MacAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[11].Value.ToString();
                st_DutsSetting[i].GuiScriptExcelFile = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[12].Value.ToString();
                //st_DutsSetting[i].FwFileName = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[12].Value.ToString();
                //st_DutsSetting[i].SSID = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[13].Value.ToString();
                //st_DutsSetting[i].IpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[1].Value.ToString();
            }

            return true;
        }

        private bool InitGuiScriptCommandSequenceRouterChamberPerformanceTest()
        {
            SetText("Check if Radius Server IP Valid?", txtRouterChamberPerformanceFunctionTestInformation);
            IPAddress ipaRadiusIP;

            if (!IPAddress.TryParse(mtbRouterChamberPerformanceConfigurationRadiusServerIpAddress.Text, out ipaRadiusIP))
            {
                SetText("Failed!", txtRouterChamberPerformanceFunctionTestInformation);
                return false;
            }

            string[] ips = ipaRadiusIP.ToString().Split('.');

            sa_GuiScriptCommandSequenceChamberPerformanceTest = new string[]
            {
                "Default Login",
                "Setting Wifi SSID 5G",
                "Setting Wifi Mode 5G",
                "Setting Wifi BandWidth 5G", 
                "Setting Wifi Channel 5G",
                "Setting Wifi Security None 5G",               
                "Setting Wifi Security 5G",
                //"Setting Wifi Security ENT KEY 5G",
                //"Setting Wifi Security Radius IP.1 5G",
                //"Setting Wifi Security Radius IP.2 5G",
                //"Setting Wifi Security Radius IP.3 5G",
                //"Setting Wifi Security Radius IP.4 5G",
                //"Setting Wifi Security Radius Port 5G",
                "",
                "",
                "",
                "",
                "",
                ""
                //"Logout"
            };

            sa_GuiScriptValueSequenceChamberPerformanceTest = new string[]
            {
                "",
                "SSID",
                "Mode",
                "BandWidth",
                "Channel", 
                "None",
                "Security",
                "Security KEY",
                ips[0],
                ips[1],
                ips[2],
                ips[3],
                nudRouterChamberPerformanceConfigurationRadiusServerPort.Value.ToString()
                //""
            };         


            //Dictionary<string, string> myDic = new Dictionary<string, string>();

            ////Dictionary<string, string>[] myDic1 = new Dictionary<string, string>[]
            ////    {

            ////    };
            //myDic.Add("Default Login", "Default Login");
            //myDic.Add("Setting Wifi SSID 5G", "Setting Wifi SSID 5G");
            //myDic.Add("Setting Wifi Channel 5G", "Setting Wifi Channel 5G");
            //myDic.Add("Setting Wifi Mode 5G", "Setting Wifi Mode 5G");
            //myDic.Add("Setting Wifi BandWidth 5G", "Setting Wifi BandWidth 5G");
            //myDic.Add("Setting Wifi Security 5G", "Setting Wifi Security 5G");
            //myDic.Add("Setting Wifi Security ENT KEY 5G", "Setting Wifi Security ENT KEY 5G");
            //myDic.Add("Setting Wifi Security Radius IP.1 5G", "Setting Wifi Security Radius IP.1 5G");
            //myDic.Add("Setting Wifi Security Radius IP.2 5G", "Setting Wifi Security Radius IP.2 5G");
            //myDic.Add("Setting Wifi Security Radius IP.3 5G", "Setting Wifi Security Radius IP.3 5G");
            //myDic.Add("Setting Wifi Security Radius IP.4 5G", "Setting Wifi Security Radius IP.4 5G");
            //myDic.Add("Setting Wifi Security Radius Port 5G", "Setting Wifi Security Radius Port 5G");
            //myDic.Add("Logout", "Logout");            

            return true;
        }

        #endregion
                
        #region Folder related

        private string createReportMainFolderRouterChamberPerformance()
        {
            string subFolder = DateTime.Now.ToString("yyyyMMdd HH-mm-ss");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

            string Path = System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder;

            return Path;
        }

        private string createSubFolderRouterChamberPerformance(string sMainFolder)
        {
            string subFolder = ((modelInfo.ModelName == "") ? "Router_" : modelInfo.ModelName + "_") + DateTime.Now.ToString("yyyyMMdd-HHmmss");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\");           

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + Path.GetFileName(sMainFolder)))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + Path.GetFileName(sMainFolder));

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + Path.GetFileName(sMainFolder) + "\\" + subFolder))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + Path.GetFileName(sMainFolder) + "\\" + subFolder);

            subFolder = System.Windows.Forms.Application.StartupPath + @"\report\" + Path.GetFileName(sMainFolder) + "\\" + subFolder;
            return subFolder;            
        }
        
        #endregion

        #region Router Gui Control

        private bool SettingRouterGuiRouterChamberPerformanceTest(string sRouterIP, WifiParameter stWifiParameter, string[,] saRounterGuiScript)
        {
            string str = string.Empty;
            bool bStatus = false;
            bool bFistTime = true;
            bool bGetValue = false;

            if (bRouterChamberPerformanceTestFTThreadRunning == false)
            {
                bRouterChamberPerformanceTestFTThreadRunning = false;
                //MessageBox.Show("Abort test", "Error");
                //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
                //threadRouterChamberPerformanceTestFT.Abort();
                return false;
                // Never go here ...
            }

            /* Initial Selenium */
            cs_BrowserChamberPerformanceTest = new CBT_SeleniumApi();
            CBT_SeleniumApi.BrowserType csbtChamberPerformanceTest = CBT_SeleniumApi.BrowserType.Chrome;

            if (!cs_BrowserChamberPerformanceTest.init(csbtChamberPerformanceTest))
            {
                str = "Initial Selenuim Failed.";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }

            str = "Start Gui Script Setting.";
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            Thread.Sleep(2000);
            LoginSetting cLoginParameter = new LoginSetting();

            CBT_SeleniumApi.GuiScriptParameter GuiScriptParameter = new CBT_SeleniumApi.GuiScriptParameter();
            string sExceptionInfo = string.Empty;
            cLoginParameter.GatewayIP = sRouterIP;
            cLoginParameter.HTTP_Port = "80";
            string sLoginURL = string.Format("http://{0}:{1}", cLoginParameter.GatewayIP, cLoginParameter.HTTP_Port);

            str = "Login URL: " + sLoginURL;
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            //string[] GuiCommand = new string[] {"Default Login", "Setting SSID", "Setting Band", "Setting Wifi Mode", "Setting Channel", "Setting Security", "Logout" };
            //string[] GuiWriteValue = new string[] {"", "Jin", stWifiParameter.Band, stWifiParameter.Mode, stWifiParameter.Channel, stWifiParameter.Security, ""};
            //string[] GuiCommand = new string[] { "Default Login", 
            //    "Setting Wifi Security 2.4G", 
            //    "Setting Wifi Security ENT KEY 5G", 
            //    "Setting Wifi Security Radius IP.1 5G", 
            //    "Setting Wifi Security Radius IP.2 5G", 
            //    "Setting Wifi Security Radius IP.3 5G", 
            //    "Setting Wifi Security Radius IP.4 5G", 
            //    "Setting Wifi Security Radius Port 5G",                 
            //    "Logout" };

            //string[] GuiCommand = new string[] { "Default Login", 
            //    "Setting Wifi Security 5G", 
            //    "Setting Wifi Security ENT KEY 5G", 
            //    "Setting Wifi Security Radius IP.1 5G", 
            //    "Setting Wifi Security Radius IP.2 5G", 
            //    "Setting Wifi Security Radius IP.3 5G", 
            //    "Setting Wifi Security Radius IP.4 5G", 
            //    "Setting Wifi Security Radius Port 5G",                 
            //    "Logout" };


            //string[] GuiWriteValue = new string[] { "", "WPA2 Enterprise", "12345678", "192", "168", "1", "100", "1234", "" };
            //str = String.Format("{0}, {1}, {2}, {3}, {4}", "", "", "", "", "");
            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            string[] GuiCommand = sa_GuiScriptCommandSequenceChamberPerformanceTest;
            string[] GuiWriteValue = sa_GuiScriptValueSequenceChamberPerformanceTest;


            //for(int CommandRow = 0; CommandRow < GuiCommand.Length; CommandRow++)
            for (int CommandRow = 0; CommandRow < GuiCommand.Length; CommandRow++)
            {
                if (bRouterChamberPerformanceTestFTThreadRunning == false)
                {
                    bRouterChamberPerformanceTestFTThreadRunning = false;
                    //MessageBox.Show("Abort test", "Error");
                    //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
                    //threadRouterChamberPerformanceTestFT.Abort();
                    return false;
                    // Never go here ...
                }

                s_CurrentGuiAccessContent = string.Empty;
                if (GuiCommand[CommandRow] == null || GuiCommand[CommandRow] == "")
                    continue;

                string sCommand = GuiCommand[CommandRow];
                //string sCommand = "Setting Wifi SSID 2.4G";
                string sWriteValue = GuiWriteValue[CommandRow];

                str = String.Format("***Run GUI Script Command: {0}", sCommand);
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                bStatus = false;
                bGetValue = false;

                for (int GuiRow = 0; GuiRow < saRounterGuiScript.GetLength(0); GuiRow++)
                {
                    if (bRouterChamberPerformanceTestFTThreadRunning == false)
                    {
                        bRouterChamberPerformanceTestFTThreadRunning = false;
                        //MessageBox.Show("Abort test", "Error");
                        //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
                        //threadRouterChamberPerformanceTestFT.Abort();
                        return false;
                        // Never go here ...
                    }

                    if (sCommand != saRounterGuiScript[GuiRow, 0]) continue;

                    //Thread.Sleep(1000);     
                    GuiScriptParameter.Index = saRounterGuiScript[GuiRow, 1];
                    GuiScriptParameter.Action = saRounterGuiScript[GuiRow, 2];
                    GuiScriptParameter.ActionName = saRounterGuiScript[GuiRow, 3];
                    GuiScriptParameter.ElementType = saRounterGuiScript[GuiRow, 4];
                    GuiScriptParameter.WriteValue = saRounterGuiScript[GuiRow, 5];
                    GuiScriptParameter.ExpectedValue = saRounterGuiScript[GuiRow, 5];
                    GuiScriptParameter.RadioBtnExpectedValueXpath = saRounterGuiScript[GuiRow, 6];
                    if (GuiScriptParameter.RadioBtnExpectedValueXpath != "" && GuiScriptParameter.RadioBtnExpectedValueXpath != null)
                    {
                        GuiScriptParameter.RadioBtnExpectedValueXpath = GuiScriptParameter.RadioBtnExpectedValueXpath.Replace('\"', '\'');
                    }
                    GuiScriptParameter.ElementXpath = saRounterGuiScript[GuiRow, 7];
                    if (GuiScriptParameter.ElementXpath != null && GuiScriptParameter.ElementXpath != "")
                    {
                        GuiScriptParameter.ElementXpath = GuiScriptParameter.ElementXpath.Replace('\"', '\'');
                    }
                    
                    GuiScriptParameter.TestTimeOut = saRounterGuiScript[GuiRow, 8];
                    GuiScriptParameter.Note = string.Empty;

                    bStatus = true;

                    //GuiScriptParameter.RadioBtnExpectedValueXpath = sdReadScriptData[GuiRow].RadioButtonExpectedValueXpath.Replace('\"', '\'');
                    //GuiScriptParameter.WriteValue = sdReadScriptData[GuiRow].WriteExpectedValue;
                    //GuiScriptParameter.ExpectedValue = sdReadScriptData[GuiRow].WriteExpectedValue;
                    //GuiScriptParameter.URL = sLoginURL + sdReadScriptData[GuiRow].WriteExpectedValue;
                    if (GuiScriptParameter.WriteValue != "" && GuiScriptParameter.WriteValue != null)
                    {
                        //if (GuiScriptParameter.WriteValue.StartsWith("#"))
                        //{
                        //    GuiScriptParameter.WriteValue = sWriteValue;
                        //    GuiScriptParameter.ExpectedValue = sWriteValue;
                        //    GuiScriptParameter.URL = sLoginURL + sWriteValue;
                        //}

                        if (GuiScriptParameter.WriteValue.IndexOf("#INPUT") >= 0)
                        {
                            GuiScriptParameter.WriteValue = sWriteValue;
                            GuiScriptParameter.ExpectedValue = sWriteValue;
                            GuiScriptParameter.URL = sLoginURL + sWriteValue;
                        }
                        
                    }

                    if (bFistTime)
                    {
                        if (!GuiScriptOpenHomePageRouterChamberPerformanceTest(sLoginURL))
                        {
                            cs_BrowserChamberPerformanceTest.Close_WebDriver();
                            cs_BrowserChamberPerformanceTest = null;
                            return false;
                        }

                        bFistTime = false;
                        Thread.Sleep(3000);
                    }

                    str = String.Format("Setting Type: {0}, Index: {1}, Action: {2}, Action Name: {3}, WriteValue: {4}", sCommand, GuiScriptParameter.Index, GuiScriptParameter.Action, GuiScriptParameter.ActionName, GuiScriptParameter.WriteValue);
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    switch (GuiScriptParameter.Action)
                    {
                        case "Set":
                            if (!GuiScriptSetFunctionRouterChamberPerformanceTest(GuiScriptParameter))
                            {
                                cs_BrowserChamberPerformanceTest.Close_WebDriver();
                                cs_BrowserChamberPerformanceTest = null;
                                return false;
                            }
                            break;
                        case "Goto":
                            if (!GuiScriptGotoFunctionRouterChamberPerformanceTest(GuiScriptParameter))
                            {
                                cs_BrowserChamberPerformanceTest.Close_WebDriver();
                                cs_BrowserChamberPerformanceTest = null;
                                return false;
                            }
                            break;
                        case "Get":
                            if (!GuiScriptGetFunctionRouterChamberPerformanceTest(GuiScriptParameter))
                            {
                                cs_BrowserChamberPerformanceTest.Close_WebDriver();
                                cs_BrowserChamberPerformanceTest = null;
                                return false;
                            }                            

                            break;
                        case "Wait":                            
                            if (!GuiScriptWaitFunctionRouterChamberPerformanceTest(GuiScriptParameter))
                            {
                                cs_BrowserChamberPerformanceTest.Close_WebDriver();
                                cs_BrowserChamberPerformanceTest = null;
                                return false;
                            }
                            break;

                        default:
                            {
                                cs_BrowserChamberPerformanceTest.Close_WebDriver();
                                cs_BrowserChamberPerformanceTest = null;
                                return false;
                            }
                    }

                    if (GuiScriptParameter.Action == "Get")
                    {
                        try
                        {
                            bGetValue = true;

                            //if (s_CurrentGuiAccessContent.Trim() == sWriteValue.Trim())
                            if (s_CurrentGuiAccessContent.Trim() == GuiScriptParameter.WriteValue.Trim())
                            { //Current Value is the same as Write Value, break GuiRow loop and run next Command
                                break;
                            }
                        }
                        catch (Exception)
                        {

                        }
                    }


                } //End of GuiRow for loop

                

                if (!bStatus)
                {
                    str = "Fail: " + sCommand + " couldn't be found in GUI Script!";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                    cs_BrowserChamberPerformanceTest.Close_WebDriver();
                    cs_BrowserChamberPerformanceTest = null;
                    return false;
                }

                if (bGetValue)
                {
                    if (s_CurrentGuiAccessContent.Trim() != GuiScriptParameter.WriteValue.Trim())
                    {
                        str = "Get Value Fail! Value is not equal to sWriteValue";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                        cs_BrowserChamberPerformanceTest.Close_WebDriver();
                        cs_BrowserChamberPerformanceTest = null;
                        return false;
                    }
                }

            } // End of Gui Command for loop

            if (cs_BrowserChamberPerformanceTest != null)
            {
                cs_BrowserChamberPerformanceTest.Close_WebDriver();
                cs_BrowserChamberPerformanceTest = null;
            }


            return true;       
        }

        //---------------------------------------//
        //-------------- Go To URL --------------//
        //---------------------------------------//
        private bool GuiScriptOpenHomePageRouterChamberPerformanceTest(string sLoginURL)
        {
            string str = string.Empty;
            str = string.Format("Open GUI Login URL: {0}", sLoginURL);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            
            s_CurrentUrlChamberPerformanceTest = sLoginURL;
            cs_BrowserChamberPerformanceTest.GoToURL(sLoginURL);
            Thread.Sleep(1000);
            return true;
        }

        //---------------------------------------//
        //-------------- Set Value --------------//
        //---------------------------------------//
        private bool GuiScriptSetFunctionRouterChamberPerformanceTest(CBT_SeleniumApi.GuiScriptParameter GuiScriptParameter)
        {
            string str = string.Empty;
            bool checkXPathResult = false;
            //CBT_SeleniumApi.GuiScriptParameter ScriptParameter = new CBT_SeleniumApi.GuiScriptParameter();

            str = string.Format("Set Parameter to GUI, Index: {0}, Action Name: {1}, Action: Set Value", GuiScriptParameter.Index, GuiScriptParameter.ActionName);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            string sWriteValue = GuiScriptParameter.WriteValue;            

            str = string.Format("Search element: {0}", GuiScriptParameter.ElementXpath);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            /* Find Element, try 1 minutes at most */
            for (int i = 0; i < 30; i++)
            {
                checkXPathResult = cs_BrowserChamberPerformanceTest.CheckXPathDisplayed(GuiScriptParameter.ElementXpath);

                if (checkXPathResult) break;
                else Thread.Sleep(2000);
            }

            if (!checkXPathResult)
            {
                str = string.Format("Failed!");
                Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }

            str = string.Format("Succeed!");
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            GuiScriptParameter.Note = string.Empty;

            Thread.Sleep(5000);
            /* Start to Set Value */
            str = string.Format("Write Value to DUT: {0}", sWriteValue);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            //if (sWriteValue == "" || sWriteValue == null)
            //{
            //    str = string.Format("Write Value is empty or null.");
            //    Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            //    return false;
            //}

            try
            {
                cs_BrowserChamberPerformanceTest.SetWebElementValue(ref GuiScriptParameter);
            }
            catch (Exception ex)
            {
                ExceptionActionUSBStorage(s_CurrentUrlChamberPerformanceTest, ref GuiScriptParameter);
                GuiScriptParameter.Note = "Set Value Error:\n" + GuiScriptParameter.Note;

                str = "Write Value Exception: " + ex.ToString();
                Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }

            return true;
        }

        private bool GuiScriptGetFunctionRouterChamberPerformanceTest(CBT_SeleniumApi.GuiScriptParameter GuiScriptParameter)
        {
            string str = string.Empty;
            bool checkXPathResult = false;
            bool bTestResult = true;
            //CBT_SeleniumApi.GuiScriptParameter ScriptParameter = new CBT_SeleniumApi.GuiScriptParameter();

            str = string.Format("Get Parameter to GUI, Index: {0}, Action Name: {1}, Action: Get Value, Expected Value: {2}", GuiScriptParameter.Index, GuiScriptParameter.ActionName, GuiScriptParameter.ExpectedValue);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            string sWriteValue = GuiScriptParameter.WriteValue;

            str = string.Format("Search element: {0}", GuiScriptParameter.ElementXpath);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            /* Find Element, try 1 minutes at most */
            for (int i = 0; i < 30; i++)
            {
                checkXPathResult = cs_BrowserChamberPerformanceTest.CheckXPathDisplayed(GuiScriptParameter.ElementXpath);

                if (checkXPathResult) break;
                else Thread.Sleep(2000);
            }

            if (!checkXPathResult)
            {
                str = string.Format("Failed!");
                Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }

            //Thread.Sleep(5000);
            str = string.Format("Succeed!");
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            GuiScriptParameter.Note = string.Empty;

            /* Start to Get Value */
            str = string.Format("Start to get value of DUT.");
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            try
            {
                bTestResult = cs_BrowserChamberPerformanceTest.GetWebElementValue(ref GuiScriptParameter);
            }
            catch (Exception ex)
            {
                ExceptionActionUSBStorage(s_CurrentUrlChamberPerformanceTest, ref GuiScriptParameter);
                GuiScriptParameter.Note = "Get Value Error:\n" + GuiScriptParameter.Note;

                str = "Get Value Exception: " + ex.ToString();
                Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }

            //if (ScriptPara.ElementType.CompareTo("RADIO_BUTTON") == 0)
            //{
            //    ScriptPara.ElementXpath = st_ReadScriptDataUSBStorage[j].RadioButtonExpectedValueXpath.Replace('\"', '\'');
            //}

            //try
            //{
            //    bTestResult = cs_BrowserChamberPerformanceTest..GetWebElementValue(ref ScriptPara); // Set Value
            //}
            //catch
            //{
            //    ExceptionActionUSBStorage(s_CurrentURLUSBStorage, ref ScriptPara);
            //    ScriptPara.Note = "Execute SubmitButton Error:\n" + ScriptPara.Note;
            //    //return false;
            //}

            str = "Read Value is " + GuiScriptParameter.GetValue;
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
            
            s_CurrentGuiAccessContent = GuiScriptParameter.GetValue;
            return true;
        }

        private bool GuiScriptWaitFunctionRouterChamberPerformanceTest(CBT_SeleniumApi.GuiScriptParameter GuiScriptParameter)
        {
            string str = string.Empty;
            int iWaitTime = 60;

            if (GuiScriptParameter.WriteValue != "" && GuiScriptParameter.WriteValue != null)
            {
                Int32.TryParse(GuiScriptParameter.WriteValue, out iWaitTime);
            }
            else
            {
                Int32.TryParse(GuiScriptParameter.TestTimeOut, out iWaitTime);
            }

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            while (true)
            {
                if (bRouterChamberPerformanceTestFTThreadRunning == false)
                {
                    bRouterChamberPerformanceTestFTThreadRunning = false;
                    //MessageBox.Show("Abort test", "Error");
                    //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
                    //threadRouterChamberPerformanceTestFT.Abort();
                    return false;
                    // Never go here ...
                }

                if (stopwatch.ElapsedMilliseconds > iWaitTime * 1000)
                {
                    break;
                }
            }

            return true;
        }

        private bool GuiScriptGotoFunctionRouterChamberPerformanceTest(CBT_SeleniumApi.GuiScriptParameter GuiScriptParameter)
        {
            string str = string.Empty;            

            str = string.Format("Goto WebPage of GUI, Index: {0}, Action Name: {1}, Action: Set Value", GuiScriptParameter.Index, GuiScriptParameter.ActionName);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            s_CurrentUrlChamberPerformanceTest = GuiScriptParameter.URL;
            Thread.Sleep(1000);
            cs_BrowserChamberPerformanceTest.GoToURL(s_CurrentUrlChamberPerformanceTest);

            return true;
        }

        

        #endregion

        #region Remote Connection
        private bool RemoteConnectionRouterChamberPerformanceTest(WifiParameter stWifiParameter, string sCommand, ref string[] sReport)
        {
            string str = string.Empty;
            string valueResponsed = string.Empty;
            string expectedValue = string.Empty;            
            int iTimeOut = 90 * 1000; // 90 Seconds
            Stopwatch stopwatch = new Stopwatch();

            TcpClient client;
            BinaryReader br;
            BinaryWriter bw;

            sCommand = "WifiConnection";
            string sCommandTemp = sCommand;
            string sSSIDTemp = stWifiParameter.SSID;
            string sSecurityTypeTemp = stWifiParameter.Security;
            string sSecurityKeyTemp = stWifiParameter.SecurityKey;
            string sAutheticationType = stWifiParameter.SecurityMode;
            if (sAutheticationType == "" || sAutheticationType == null)
                sAutheticationType = "AES";
            string sSecurityIndex = sAutheticationType;

            if (sSSIDTemp == null || sSSIDTemp == "" || sSecurityTypeTemp == null || sSecurityTypeTemp == "")                
            {
                sReport[0] = "Error";
                sReport[1] = "SSID or Security Type is Empty.";
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }

            //string[] s = stWifiParameter.Security.Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);

            //sSecurityTypeTemp = s[0];

            //if (s.Length >= 2)
            //{
            //    if (!Int32.TryParse(s[1], out iSecurityIndex))
            //    {
            //        sReport[0] = "Error";
            //        sReport[1] = "SSID or Security Type is Empty or Invalid.";
            //        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
            //        return false;
            //    }
            //    else
            //    {

            //    }
            //}
           
            str = String.Format("Run Remote Control Function:");
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

            try
            {               
                switch (sCommand)
                {                        
                    case "WifiConnection":
                        {
                            if (sSecurityTypeTemp.ToLower() == "wep")
                            {                                
                                sCommandTemp = String.Format("WifiConnection::{0}::{1}::{2}::{3}::{4}", sSSIDTemp, sSecurityTypeTemp, sSecurityIndex, sAutheticationType, sSecurityKeyTemp);
                            }
                            else
                            {
                                sCommandTemp = String.Format("WifiConnection::{0}::{1}::{2}::{3}::{4}", sSSIDTemp, sSecurityTypeTemp, sAutheticationType, "1", sSecurityKeyTemp);
                            }
                            //    string sAuthetication = sa_CableIntegrationProvisionDataSets[0, 6];
                            //    string sEncryption = sa_CableIntegrationProvisionDataSets[0, 7];
                            //    string sKeyIndex = sa_CableIntegrationProvisionDataSets[0, 8];
                            //    sCommandTemp = String.Format("WifiConnection::{0}::{1}::{2}::{3}::{4}", sSSID, sAuthetication, sEncryption, sKeyIndex, sKey);
                                str = String.Format("Remote Command is {0}", sCommandTemp);
                                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                            }
                            break;                        
                    }                
            }
            catch (Exception ex)
            {
                sReport[0] = "Error";
                sReport[1] = "Setting Remote Control Parameter Error : " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }

            try
            {
                client = new TcpClient(mtbRouterChamberPerformanceConfigurationRemoteControlServerIpAddress.Text, Convert.ToInt32(nudRouterChamberPerformanceConfigurationRemoteControlServerPort.Value));

                str = "Connecte to server succeed!";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
            }
            catch (Exception ex)
            {
                sReport[0] = "Error";
                sReport[1] = "Connect to Server Failed: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }

            try
            {
                NetworkStream clientStream = client.GetStream();
                bw = new BinaryWriter(clientStream);
                //string message = "PING::168.95.1.1";
                string message = sCommandTemp;
                bw.Write(message);

                while (true)
                {
                    if (bRouterChamberPerformanceTestFTThreadRunning == false)
                    {
                        bRouterChamberPerformanceTestFTThreadRunning = false;
                        //MessageBox.Show("Abort test", "Error");
                        //this.Invoke(new showCableSwIntegrationTestGUIDelegate(ToggleCableSwIntegrationFunctionTestGUI));
                        //threadCableSwIntegrationTestFT.Abort();
                        return false;
                        // Never go here ...
                    }
                                        
                    br = new BinaryReader(clientStream);
                    string receive = null;

                    //var count = br.BaseStream.Length / sizeof(int);
                    //for (var i = 0; i < count; i++)
                    //{
                    //    int v = br.ReadInt32();

                    //    str = v.ToString();
                    //    this.Invoke(new SetControllerContent(SetTextLine), new object[] { str, txtAteRemoteControlClientInformation });
                    //    receive += str;
                    //}


                    receive = br.ReadString();
                    Invoke(new SetTextCallBackT(SetText), new object[] { receive, txtRouterChamberPerformanceFunctionTestInformation });

                    if (receive.IndexOf("RESULT") >= 0)
                    {
                        str = receive;
                        client.Close();
                        break;
                    }
                    //txtAteRemoteControlClientInformation.AppendText(receive + "\r\n");
                    //textBox.Dispatcher.Invoke(() => textBox.Text += receive + "\r\n");


                    //bw = new BinaryWriter(clientStream);
                    //string message = "PING 168.95.1.1";
                    //bw.Write(message);
                    //textBox.Text += message.Text + "\r\n";

                    //}
                    //catch (Exception ex)
                    //{
                    //    MessageBox.Show("Client receive Fail: " + ex.ToString());
                    //}
                }
            }
            catch (Exception ex)
            {
                sReport[0] = "Error";
                sReport[1] = "Client receive Fail: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                if (client != null) client.Close();
                return false;
            }          

            /* Check Result */
            int index = str.IndexOf("RESULT");
            if (index < 0)
            {
                sReport[0] = "Error";
                sReport[1] = "Remote Access Error";
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }

            string[] sResult = str.Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
            if (sResult.Length < 2)
            {
                sReport[0] = "Error";
                sReport[1] = "Remote Access Error";
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }            

            switch (sCommand)
            {
                case "WebAccess":
                    {
                        sReport[0] = sResult[1];
                        sReport[1] = sResult[2];
                        return true;
                    }
                    break;
                case "PING":
                    {                        
                        sReport[0] = sResult[1];
                        sReport[1] = sResult[2];
                        return true;                       
                    }
                    break;
                case "IPCONFIG":
                    {
                        //if (sa_CableIntegrationProvisionDataSets[0, 5] == null || sa_CableIntegrationProvisionDataSets[0, 5] == "")
                        //{
                        //    sReport = "Error";
                        //    sComment = "Expected Value Error";
                        //    Invoke(new SetTextCallBackT(SetText), new object[] { sComment, txtRouterChamberPerformanceFunctionTestInformation });
                        //    return false;
                        //}

                        if (sCommandTemp.IndexOf("Release") >= 0)
                        {
                            sReport[0] = sResult[1];
                            Invoke(new SetTextCallBackT(SetText), new object[] { sReport, txtRouterChamberPerformanceFunctionTestInformation });
                        }
                        else
                        {
                            bool bHasFail = false;
                            //string sExpectedValue = sa_CableIntegrationProvisionDataSets[0, 7].Trim();
                            //string[] sa = sExpectedValue.Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
                            //for (int i = 0; i < sa.Length; i++)
                            //{
                            //    if (str.IndexOf(sa[i]) < 0)
                            //        bHasFail = true;
                            //}

                            //if (str.IndexOf(sExpectedValue) >= 0)
                            if (!bHasFail)
                            {
                                sReport[0] = "PASS";
                                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                            }
                            else
                            {
                                sReport[0] = "FAIL";
                                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterChamberPerformanceFunctionTestInformation });
                            }
                        }                       
                    }
                    break;

                case "WifiConnection":
                    {
                        sReport[0] = sResult[1];
                        if (sResult.Length >= 3)
                        {
                            if (sResult[2].IndexOf("Exception") >= 0)
                            {
                                sReport[1] = "Remote Control Exception";
                            }
                            else
                            {
                                sReport[1] = sResult[2];
                            }
                        }                       
                    }
                    break;
                

                default:
                    {
                        str = "Unknown Remote Command";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

                        sReport[0] = "Error";
                        sReport[1] = "Unknown Remote Command";
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterChamberPerformanceFunctionTestInformation });
                        return false;
                    }

            }

            if (sReport[0] == "PASS")
                return true;
            else
                return false;   
        }
        
        #endregion

        #region Chariot Function

        private bool ChariotFunctionRouterChamberPerformanceTest(string sTstFile, string sOutputFile, ref string sResultValue, ref string sComment)
        {
            string str = string.Empty;
            string sInputFile = string.Empty;
            sResultValue = "";

            if (!File.Exists(sTstFile))
            {
                str = "File doesn't exist:" + sTstFile;
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                str = "Skip the item.";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                sComment = "Tst File doesn't exist!";
                return false;
            }
            
            sInputFile = sTstFile;
            Path_runtst = Path.GetDirectoryName(txtRouterChamberPerformanceConfigurationChariotFolder.Text) + "\\runtst.exe";
            Path_fmttst = Path.GetDirectoryName(txtRouterChamberPerformanceConfigurationChariotFolder.Text) + "\\fmttst.exe";
            
            
            try
            {
                RunChariotConsole(Path_runtst, Path_fmttst, sInputFile, sOutputFile, txtRouterChamberPerformanceFunctionTestInformation);
                string csvFile = sOutputFile + ".csv";
                sResultValue = ThroughputValue(csvFile);
            }
            catch (Exception ex)
            {
                sResultValue = "";
                str = "Run Chariot Exception: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                return false;
            }                 

            return true;
        }

        #endregion

        #region PingFunction

        private bool QuickPingChamberPerformanceTest(string ip, int iTimeout = 3000)
        {
            /* Ping if the deivce available*/
            bool pingStatus = false;
            //SetText("Ping Device...", txtCableSwIntegrationFunctionTestInformation);
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Reset();
            stopWatch.Stop();
            stopWatch.Start();

            while (true)
            {
                if (bRouterChamberPerformanceTestFTThreadRunning == false)
                {
                    bRouterChamberPerformanceTestFTThreadRunning = false;
                    //MessageBox.Show("Abort test", "Error");
                    //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
                    //threadRouterChamberPerformanceTestFT.Abort();
                    return false;
                    // Never go here ...
                }

                if (stopWatch.ElapsedMilliseconds > iTimeout)
                {
                    break;
                }

                if (PingClient(ip, 1000))
                {
                    pingStatus = true;
                    break;
                }

                Thread.Sleep(1000);
            }

            if (!pingStatus)
            {
                //SetText("Ping Failed!!", txtCableSwIntegrationFunctionTestInformation);
                //MessageBox.Show("Ping Failed, Abort Test!!", "Error");
                return false;
            }

            //SetText("Ping Succeed!!", txtCableSwIntegrationFunctionTestInformation);
            return true;
        }

        #endregion

        private string createReportExcelFileChamberPerformanceTest(string subFolder,string sModelName, int Loop, string FileName)
        {
            /* string PathFile = @"\report\SUBFOLDER\MODEL_SN_SW_HW_QAMTYPE_LEVELdBmV_(POS)_DATE_(LOOP).xlsx";   */
            string fileNameWoExt = Path.GetFileNameWithoutExtension(FileName);

            //string PathFile = @"\report\SUBFOLDER\FileName_SKU_Report_DATE_(LOOP).xlsx";
            string PathFile = @"SUBFOLDER\Model_FileName_Report_DATE_(LOOP).xlsx";

            PathFile = PathFile.Replace("SUBFOLDER", subFolder);
            PathFile = PathFile.Replace("Model", sModelName);
            PathFile = PathFile.Replace("FileName", fileNameWoExt);
            PathFile = PathFile.Replace("DATE", DateTime.Now.ToString("yyyyMMdd HH-mm"));
            PathFile = PathFile.Replace("LOOP", Loop.ToString());

            return PathFile;
        }

        #region Send Report and Alarm

        private bool WriteFailMsgToExcelAndSendEamilChamberPerformanceTest(string sFailMsg, string sLogFile)
        {
            string str = string.Empty;

            try
            {
                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = sFailMsg;

                /* Write Console Log File as HyperLink in Excel report */
                xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3];
                xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sLogFile, Type.Missing, "ConsoleLog", "ConsoleLog");
            }
            catch (Exception ex)
            {
                str = "Write Data to Excel File Exception: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
            }

            try
            {
                /* Save end time and close the Excel object */
                xls_excelAppRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                xls_excelWorkBookRouterChamberPerformanceTest.Save();
                /* Close excel application when finish one job(power level) */
                closeExcelRouterChamberPerformanceTest();
                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                //str = "SavWrite Data to Excel File Exception: " + ex.ToString();
                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
            }

            //寄出報告
            try
            {
                str = "Router Chamber Performance Test Failed: " + modelInfo.ModelName + " " + sFailMsg;                
                SendReportByGmailWithoutFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, sFailMsg);
            }
            catch (Exception ex)
            {
                str = "Send Mail Failed: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                return false;                    
            }

            //string sLogFile = "";
            File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
            //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

            return true;
        }

        private bool SendLineAlarmChamberPerformanceTest(string sFailMsg)
        {
            string str = string.Empty;

            try
            {
                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = sFailMsg;

                /* Write Console Log File as HyperLink in Excel report */
                xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3];
                //xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
            }
            catch (Exception ex)
            {
                str = "Write Data to Excel File Exception: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
            }
            //    /* Save end time and close the Excel object */
            //    xls_excelAppCableSwIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            //    xls_excelWorkBookCableSwIntegrationTest.Save();
            //    /* Close excel application when finish one job(power level) */
            //    closeExcelCableSwIntegrationTest();
            //    Thread.Sleep(3000);
            //    //寄出報告
            //    try
            //    {
            //        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Read Excel File Failed", "Read Excel File Failed!!!", savePath);
            //    }
            //    catch (Exception ex)
            //    {
            //        str = "Send Mail Failed: " + ex.ToString();
            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
            //    }
            //    return false;
            //}

            return true;
        }




        #endregion 

        /*============================================================================================*/
        /*================================= Delegate Function Area   =================================*/
        /*============================================================================================*/
        #region

        private void ToggleRouterChamberPerformanceFunctionTestGUI()
        {
            ToggleRouterChamberPerformanceFunctionTestController(true);
            Debug.WriteLine("Toggle");
        }

        /* Disable/Enable controller */
        private void ToggleRouterChamberPerformanceFunctionTestController(bool Toggle)
        {
            /* Model Section */
            txtRouterChamberPerformanceFunctionTestName.Enabled = Toggle;
            txtRouterChamberPerformanceFunctionTestSwVersion.Enabled = Toggle;
            txtRouterChamberPerformanceFunctionTestSerialNumber.Enabled = Toggle;
            txtRouterChamberPerformanceFunctionTestHwVersion.Enabled = Toggle;
            btnRouterChamberPerformanceFunctionTestSaveLog.Enabled = Toggle;            

            /* Time setting */
            nudRouterChamberPerformanceFunctionTestPingTimeout.Enabled = Toggle;
            nudRouterChamberPerformanceFunctionTestConditionTimeout.Enabled = Toggle;
            nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Enabled = Toggle;

            /* Loop Times */
            chkRouterChamberPerformanceFunctionTestScheduleOnOff.Enabled = Toggle;
            nudRouterChamberPerformanceFunctionTestTimes.Enabled = Toggle;            

            /* Duts Setting Section */
            btnRouterChamberPerformanceDutsSettingAddSetting.Enabled = Toggle;
            btnRouterChamberPerformanceDutsSettingEditSetting.Enabled = Toggle;
            btnRouterChamberPerformanceDutsSettingSaveSetting.Enabled = Toggle;
            btnRouterChamberPerformanceDutsSettingLoadSetting.Enabled = Toggle;
            btnRouterChamberPerformanceDutsSettingMoveUp.Enabled = Toggle;
            btnRouterChamberPerformanceDutsSettingMoveDown.Enabled = Toggle;

            nudRouterChamberPerformanceDutsSettingIndex.Enabled = Toggle;
            txtRouterChamberPerformanceDutsSettingModelName.Enabled = Toggle;
            txtRouterChamberPerformanceDutsSettingSerialNumber.Enabled = Toggle;
            txtRouterChamberPerformanceDutsSettingSwVersion.Enabled = Toggle;
            txtRouterChamberPerformanceDutsSettingHwVersion.Enabled = Toggle;
            mtbRouterChamberPerformanceDutsSettingIpAddress.Enabled = Toggle;
            mtbRouterChamberPerformanceDutsSettingMacAddress.Enabled = Toggle;
            mtbRouterChamberPerformanceDutsSettingPcIpAddress.Enabled = Toggle;
            cboxRouterChamberPerformanceDutsSettingComPort.Enabled = Toggle; 
            nudRouterChamberPerformanceDutsSettingwitchPort.Enabled = Toggle;
            txtRouterChamberPerformanceDutsSetting24gSsid.Enabled = Toggle;
            txtRouterChamberPerformanceDutsSettingGuiScripExcelFileName.Enabled = Toggle;
            txtRouterChamberPerformanceDutsSetting5gSsid.Enabled = Toggle;         

            ///* SNMP Setting */
            //rbtnRouterChamberPerformanceConfigurationSnmpV1.Enabled = Toggle;
            //rbtnRouterChamberPerformanceConfigurationSnmpV2.Enabled = Toggle;
            //rbtnRouterChamberPerformanceConfigurationSnmpV3.Enabled = Toggle;
            //nudRouterChamberPerformanceConfigurationSnmpPort.Enabled = Toggle;
            //nudRouterChamberPerformanceConfigurationTrapPort.Enabled = Toggle;
            //dudRouterChamberPerformanceConfigurationSnmpReadCommunity.Enabled = Toggle;
            //dudRouterChamberPerformanceConfigurationSnmpWriteCommunity.Enabled = Toggle;           

            /*Other Setting */
            //dtpRouterChamberPerformanceConfigurationFtpTestPeriodStartDay.Enabled = Toggle;
            //dtpRouterChamberPerformanceConfigurationFtpTestPeriodEndDay.Enabled = Toggle;
            //txtRouterChamberPerformanceConfigurationRdFtpServerPath.Enabled = Toggle;
            //txtRouterChamberPerformanceConfigurationRdFtpFwFileName.Enabled = Toggle;
            //txtRouterChamberPerformanceConfigurationNfsServerPath.Enabled = Toggle;

            /* Email Setting */
            chkRouterChamberPerformanceConfigurationSendReport.Enabled = Toggle;
            txtRouterChamberPerformanceConfigurationReportEmailSendTo.Enabled = Toggle;
            txtRouterChamberPerformanceConfigurationReportEmailSenderGmailAccount.Enabled = Toggle;
            txtRouterChamberPerformanceConfigurationReportEmailSenderGmailPassword.Enabled = Toggle;
            
            /* Line Message Setting */
            chkRouterChamberPerformanceConfigurationLineMessage.Enabled = Toggle;
            txtRouterChamberPerformanceConfigurationLineMessageXmlFileName.Enabled = Toggle;
            btnRouterChamberPerformanceConfigurationLineMessageXmlFileName.Enabled = Toggle;            

            /* Swtich Setting */
            chkRouterChamberPerformanceConfigurationUseSwitch.Enabled = Toggle;
            txtRouterChamberPerformanceConfigurationSwitchUsername.Enabled = Toggle;
            txtRouterChamberPerformanceConfigurationSwitchPassword.Enabled = Toggle;

            /* Others Setting */
            mtbRouterChamberPerformanceConfigurationRemoteControlServerIpAddress.Enabled = Toggle;
            nudRouterChamberPerformanceConfigurationRemoteControlServerPort.Enabled = Toggle;
            mtbRouterChamberPerformanceConfigurationClientCardIpAddress.Enabled = Toggle;
            txtRouterChamberPerformanceConfigurationTestConditionExcelFileName.Enabled = Toggle;
            btnRouterChamberPerformanceConfigurationTestConditionExcelFile.Enabled = Toggle;
            txtRouterChamberPerformanceConfigurationChariotFolder.Enabled = Toggle;
            btnRouterChamberPerformanceConfigurationChariotFolder.Enabled = Toggle;

            /* Radius Server Section */
            mtbRouterChamberPerformanceConfigurationRadiusServerIpAddress.Enabled = Toggle;
            nudRouterChamberPerformanceConfigurationRadiusServerPort.Enabled = Toggle;

            btnRouterChamberPerformanceFunctionTestSaveLog.Enabled = Toggle;
            btnRouterChamberPerformanceFunctionTestRun.Text = Toggle ? "Run" : "Stop";
            Debug.WriteLine(btnRouterChamberPerformanceFunctionTestRun.Text);

            /* Disable/Enable "Items" and "Setup" menu item */
            //functionToolStripMenuItem.Enabled = Toggle;
            setupToolStripMenuItem.Enabled = Toggle;

            /* Show total elapsed testing time */
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

            if (bRouterChamberPerformanceTestFTThreadRunning == false)
            {
                /* Save excel data */
                if (xls_excelWorkBookRouterChamberPerformanceTest != null)
                {
                    try
                    {
                        /* Save end time and close the Excel object */
                        xls_excelAppRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                        xls_excelWorkBookRouterChamberPerformanceTest.Save();
                        closeExcelRouterChamberPerformanceTest();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Save Excel Error: " + ex.ToString());
                    }
                }

                if (cs_BrowserChamberPerformanceTest != null)
                    cs_BrowserChamberPerformanceTest.Close_WebDriver();

                MessageBox.Show(this, "Test complete !!!", "Information", MessageBoxButtons.OK);
            }
        }

        #endregion

        /*============================================================================================*/
        /*=================================== Excel Function Area   ==================================*/
        /*============================================================================================*/
        #region
        ///<summary>
        ///The following function call processes the related of Excel application
        ///</summary>
        private void initialExcelRouterChamberPerformanceTest(string savePath)
        {
            ////Does Excel.exe process is running?

#if DO_EXIST_EXCEL
                    bool flag = false ;

                    foreach(Process item in Process.GetProcesses())
                    {
                        if(item.ProcessName == "EXCEL")
                        {
                            flag= true ;
                            break;
                        }
                    }

                    for(int i=0; i<3; i++)
                    {
                        if(!flag) /* Check exist process */
                        {
                            xls_excelAppRouterChamberPerformanceTest= new Excel.Application();
                        }
                        else /* Use exist process */
                        {
                            object obj = Marshal.GetActiveObject("Excel.Application");
                            xls_excelAppRouterChamberPerformanceTest = obj as Excel.Application;                 
                        }

                        if(xls_excelAppRouterChamberPerformanceTest !=null)
                            break;

                        Thread.Sleep(3000*(i+1));
                    }
#else
            xls_excelAppRouterChamberPerformanceTest = new Excel.Application();
#endif
            if (xls_excelAppRouterChamberPerformanceTest == null)
            {
                bRouterChamberPerformanceTestFTThreadRunning = false;
                MessageBox.Show("Excel process could not be forked !!!", "Error");
                //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
                //threadRouterChamberPerformanceTestFT.Abort();
                return;
                //Never go here
            }

            //Maxmimize excel windows
            xls_excelAppRouterChamberPerformanceTest.WindowState = Excel.XlWindowState.xlMaximized;

            /* Set Excel visible */
            xls_excelAppRouterChamberPerformanceTest.Visible = true;

            /* do not show alert */
            xls_excelAppRouterChamberPerformanceTest.DisplayAlerts = false;

            xls_excelAppRouterChamberPerformanceTest.UserControl = true;
            xls_excelAppRouterChamberPerformanceTest.Interactive = false;


            /* Set font and font size attributes */
            xls_excelAppRouterChamberPerformanceTest.StandardFont = "Arial";
            xls_excelAppRouterChamberPerformanceTest.StandardFontSize = 10;

            xls_excelWorkBookRouterChamberPerformanceTest = xls_excelAppRouterChamberPerformanceTest.Workbooks.Add(misValue); /* This method is used to open an Excel workbook by passing the file path as a parameter to this method. */

            xls_excelWorkSheetRouterChamberPerformanceTest = (Excel.Worksheet)xls_excelWorkBookRouterChamberPerformanceTest.Sheets[1]; /* By default, every workbook created has three worksheets */
            xls_excelWorkSheetRouterChamberPerformanceTest.Name = "Results";
            createTitleExcelRouterChamberPerformanceTest();

            try
            {
                xls_excelWorkBookRouterChamberPerformanceTest.SaveAs(savePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp);
            }
        }

        private void createTitleExcelRouterChamberPerformanceTest()
        {
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 2] = "CyberATE " + modelInfo.ModelName + " Router Chamber Performance Test Report";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[4, 2] = "Test";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[5, 2] = "Station";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[6, 2] = "Start Time";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 2] = "End Time";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[8, 2] = "Model";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[9, 2] = "Serial";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 2] = "SW Version";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[11, 2] = "HW Version";

            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[4, 3] = modelInfo.ModelName + " Router Chamber Performance Test";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[5, 3] = "CyberTAN Router-" + modelInfo.ModelName;

            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[6, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[8, 3] = modelInfo.ModelName;
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[9, 3] = modelInfo.SN;
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 3] = modelInfo.SwVersion;
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[11, 3] = modelInfo.HwVersion;

            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[6, 3], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[11, 3]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            /* Set cells width */
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 1], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 1]].ColumnWidth = 2;
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 2], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 2]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 3], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 3]].ColumnWidth = 15;
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 4], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 4]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 5]].ColumnWidth = 15;
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 6], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 6]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 7], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 7]].ColumnWidth = 15;
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 8], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 8]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 9], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 9]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 10], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 10]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[2, 11], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 11]].ColumnWidth = 15; /* column B */

            xls_excelAppRouterChamberPerformanceTest.Cells[3, 8] = "Option";
            xls_excelAppRouterChamberPerformanceTest.Cells[3, 9] = "Value";

            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelAppRouterChamberPerformanceTest.Cells[3, 8], xls_excelAppRouterChamberPerformanceTest.Cells[8, 9]].Borders.LineStyle = 1;

            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelAppRouterChamberPerformanceTest.Cells[3, 8], xls_excelAppRouterChamberPerformanceTest.Cells[3, 9]].Font.Underline = true;
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelAppRouterChamberPerformanceTest.Cells[3, 8], xls_excelAppRouterChamberPerformanceTest.Cells[3, 9]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelAppRouterChamberPerformanceTest.Cells[3, 8], xls_excelAppRouterChamberPerformanceTest.Cells[3, 9]].Font.FontStyle = "Bold";

            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[4, 8] = "Loop";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[5, 8] = "Gui Script File";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[6, 8] = "Tx Tst";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 8] = "Rx Tst";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[8, 8] = "Bi Tst";

            /* Fill the Title */
            int rowCount = 15;
            //xls_excelAppRouterChamberPerformanceTest.Cells[rowCount, 2] = "Index";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 2] = "Index";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 3] = "Function Name";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 4] = "Name";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 5] = "Band";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 6] = "Mode";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 7] = "Channel";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 8] = "BandWidth";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 9] = "Security";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 10] = "Security Key";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 11] = "Tx(Mbps)";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 12] = "Rx(Mbps)";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 13] = "Bi(Mbps)";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, i_CommentColumnChamberPerformanceTest] = "Comment";
            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, i_LogColumnChamberPerformanceTest] = "Log File";
            //xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 10] = "Log File";
            //xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 11] = "Test Result";
            //xls_excelWorkSheetRouterChamberPerformanceTest.Cells[rowCount, 12] = "Comment";

            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[14, 2], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[14, 10]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }      

        private void saveExcelRouterChamberPerformanceTest(string savePath)
        {
            try
            {
                xls_excelWorkBookRouterChamberPerformanceTest.SaveAs(savePath, misValue, misValue, misValue,
                    misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue,
                    misValue, misValue, misValue);
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp);
            }
        }       

        private void closeExcelRouterChamberPerformanceTest()
        {
            /* Turn on interactive mode */
            xls_excelAppRouterChamberPerformanceTest.Interactive = true;
            xls_excelWorkBookRouterChamberPerformanceTest.Close();
            xls_excelAppRouterChamberPerformanceTest.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xls_excelWorkSheetRouterChamberPerformanceTest);
            xls_excelWorkSheetRouterChamberPerformanceTest = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xls_excelWorkBookRouterChamberPerformanceTest);
            xls_excelWorkBookRouterChamberPerformanceTest = null;

            if (xls_excelRangeRouterChamberPerformanceTest != null)
            {
                releaseObject_RouterChamberPerformanceTest(xls_excelRangeRouterChamberPerformanceTest);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xls_excelRangeRouterChamberPerformanceTest);
                xls_excelRangeRouterChamberPerformanceTest = null;
            }
            releaseObject_RouterChamberPerformanceTest(xls_excelAppRouterChamberPerformanceTest);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xls_excelAppRouterChamberPerformanceTest);
            xls_excelAppRouterChamberPerformanceTest = null;

            GC.Collect();
        }

        private void releaseObject_RouterChamberPerformanceTest(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        #endregion

        /*============================================================================================*/
        /*=================================== XML Function Area   ====================================*/
        /*============================================================================================*/
        #region

        ///
        /// XML parsing
        ///
        public void writeXmlDefaultRouterChamberPerformanceTest(string FileName)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;

            XmlWriter writer = XmlWriter.Create(FileName, settings);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file is generated by the program.");
            writer.WriteStartElement("RouterATE");
            writer.WriteAttributeString("Item", "Router Chamber Performance Test");

            ///
            /// Write Function Test settings
            /// 
            writer.WriteStartElement("FunctionTest");

            // Model section
            writer.WriteStartElement("Model");
            writer.WriteElementString("Name", "CGA2121");
            writer.WriteElementString("SN", "0123456789");
            writer.WriteElementString("SWVer", "Diag. 1.0");
            writer.WriteElementString("HWVer", "ES 1.0");
            writer.WriteEndElement();

            /* Time Setting*/
            writer.WriteStartElement("TimeSetting");
            writer.WriteElementString("PingTimeout", "60");
            writer.WriteElementString("ConditionTimeout", "40");
            writer.WriteElementString("DutRebootTimeout", "600");
            writer.WriteEndElement();

            /* Loop Times */
            writer.WriteStartElement("LoopSetting");
            writer.WriteElementString("ScheduleOnOff", "Off");
            writer.WriteElementString("ScheduleTimes", "1");
            writer.WriteEndElement();

            writer.WriteEndElement();
            // End of write Function Test

            ///
            /// Write Configuration settings
            /// 
            writer.WriteStartElement("Configuration");
           
            //Email Section
            writer.WriteStartElement("EmailSetting");
            writer.WriteElementString("SendReport", "Y");
            writer.WriteElementString("ReportSenderAccount", "cbt.sqa5503@gmail.com");
            writer.WriteElementString("ReportSenderPassword", "9gu112a8");
            writer.WriteElementString("ReportSendTo", "jin.wang@cybertan.com.tw;jin.wang@cybertan.com.tw");
            writer.WriteEndElement();

            // Line Section
            writer.WriteStartElement("LineSetting");
            writer.WriteElementString("LineMessage", "Y");
            writer.WriteElementString("LineGroupXmlFile", "");        
            writer.WriteEndElement();

            // Switch Section
            writer.WriteStartElement("SwitchSetting");
            writer.WriteElementString("UseSwitch", "N");
            writer.WriteElementString("Username", "cisco");
            writer.WriteElementString("Password", "sqa11063");
            writer.WriteEndElement();

            // Other Setting Section
            writer.WriteStartElement("OtherSetting");
            writer.WriteElementString("RemoteServerIp", "192.168.0.12");
            writer.WriteElementString("RemoteServerPort", "7890");
            writer.WriteElementString("ClientCardIp", "192.168.1.100");
            writer.WriteElementString("ChariotPath", @"C:\Program Files\Ixia\IxChariot\IxChariot.exe");
            writer.WriteElementString("TestConditionFileName", "");
            writer.WriteEndElement();

            // Radius Server Section
            writer.WriteStartElement("RadiusServerSetting");
            writer.WriteElementString("RadiusServerIp", "192.168.1.100");
            writer.WriteElementString("RadiusServerPort", "1234");            
            writer.WriteEndElement();

            writer.WriteEndElement();
            // End of write Configuration

            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public void writeXmlRouterChamberPerformanceTest(string FileName)
        {
            // ToDo: Needs to verify which test item is to be selected..

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;

            XmlWriter writer = XmlWriter.Create(FileName, settings);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file is generated by the program.");
            writer.WriteStartElement("RouterATE");
            writer.WriteAttributeString("Item", "Router Chamber Performance Test");

            ///
            /// Write Function Test settings
            /// 
            writer.WriteStartElement("FunctionTest");

            // Model section
            writer.WriteStartElement("Model");
            writer.WriteElementString("Name", txtRouterChamberPerformanceFunctionTestName.Text);
            writer.WriteElementString("SN", txtRouterChamberPerformanceFunctionTestSerialNumber.Text);
            writer.WriteElementString("SWVer", txtRouterChamberPerformanceFunctionTestSwVersion.Text);
            writer.WriteElementString("HWVer", txtRouterChamberPerformanceFunctionTestHwVersion.Text);
            writer.WriteEndElement();

            /* Time Setting*/
            writer.WriteStartElement("TimeSetting");
            writer.WriteElementString("PingTimeout", nudRouterChamberPerformanceFunctionTestPingTimeout.Value.ToString());
            writer.WriteElementString("ConditionTimeout", nudRouterChamberPerformanceFunctionTestConditionTimeout.Value.ToString());
            writer.WriteElementString("DutRebootTimeout", nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString());
            writer.WriteEndElement();

            /* Loop Times */
            writer.WriteStartElement("LoopSetting");
            writer.WriteElementString("ScheduleOnOff", chkRouterChamberPerformanceFunctionTestScheduleOnOff.Checked ? "On" : "Off");
            writer.WriteElementString("ScheduleTimes", nudRouterChamberPerformanceFunctionTestTimes.Value.ToString());
            writer.WriteEndElement();

            writer.WriteEndElement();
            // End of write Function Test

            ///
            /// Write Configuration settings
            /// 
            writer.WriteStartElement("Configuration");

            //Email Section
            writer.WriteStartElement("EmailSetting");
            writer.WriteElementString("SendReport", chkRouterChamberPerformanceConfigurationSendReport.Checked? "Y":"N");
            writer.WriteElementString("ReportSenderAccount", txtRouterChamberPerformanceConfigurationReportEmailSenderGmailAccount.Text);
            writer.WriteElementString("ReportSenderPassword", txtRouterChamberPerformanceConfigurationReportEmailSenderGmailPassword.Text);
            writer.WriteElementString("ReportSendTo", txtRouterChamberPerformanceConfigurationReportEmailSendTo.Text);
            writer.WriteEndElement();

            // Line Section
            writer.WriteStartElement("LineSetting");
            writer.WriteElementString("LineMessage", chkRouterChamberPerformanceConfigurationLineMessage.Checked ? "Y" : "N");
            writer.WriteElementString("LineGroupXmlFile", txtRouterChamberPerformanceConfigurationLineMessageXmlFileName.Text);
            writer.WriteEndElement();           

            //Switch Section
            writer.WriteStartElement("SwitchSetting");
            writer.WriteElementString("UseSwitch", chkRouterChamberPerformanceConfigurationUseSwitch.Checked ? "Y" : "N");
            writer.WriteElementString("Username", txtRouterChamberPerformanceConfigurationSwitchUsername.Text);
            writer.WriteElementString("Password", txtRouterChamberPerformanceConfigurationSwitchPassword.Text);
            writer.WriteEndElement();                    

            //Other Setting
            writer.WriteStartElement("OtherSetting");
            writer.WriteElementString("RemoteServerIp", mtbRouterChamberPerformanceConfigurationRemoteControlServerIpAddress.Text);
            writer.WriteElementString("RemoteServerPort", nudRouterChamberPerformanceConfigurationRemoteControlServerPort.Value.ToString());
            writer.WriteElementString("ClientCardIp", mtbRouterChamberPerformanceConfigurationClientCardIpAddress.Text);
            writer.WriteElementString("ChariotPath", txtRouterChamberPerformanceConfigurationChariotFolder.Text);
            writer.WriteElementString("TestConditionFileName", txtRouterChamberPerformanceConfigurationTestConditionExcelFileName.Text);
            writer.WriteEndElement();

            // Radius Server Section
            writer.WriteStartElement("RadiusServerSetting");
            writer.WriteElementString("RadiusServerIp", mtbRouterChamberPerformanceConfigurationRadiusServerIpAddress.Text);
            writer.WriteElementString("RadiusServerPort", nudRouterChamberPerformanceConfigurationRadiusServerPort.Value.ToString());
            writer.WriteEndElement();

            writer.WriteEndElement();
            // End of write Configuration          

            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public bool readXmlRouterChamberPerformanceTest(string FileName)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(FileName);

            XmlNode nodeXml = doc.SelectSingleNode("RouterATE");
            if (nodeXml == null)
                return false;
            XmlElement element = (XmlElement)nodeXml;
            string strID = element.GetAttribute("Item");
            Debug.WriteLine(strID);
            if (strID.CompareTo("Router Chamber Performance Test") != 0)
            {
                MessageBox.Show("This XML file is incorrect.", "Error");
                return false;
            }

            ///
            /// Read Function Test configuration settings
            ///
            // Model
            XmlNode nodeFunctionTestModel = doc.SelectSingleNode("/RouterATE/FunctionTest/Model");
            try
            {
                string Name = nodeFunctionTestModel.SelectSingleNode("Name").InnerText;
                string SN = nodeFunctionTestModel.SelectSingleNode("SN").InnerText;
                string SWVer = nodeFunctionTestModel.SelectSingleNode("SWVer").InnerText;
                string HWVer = nodeFunctionTestModel.SelectSingleNode("HWVer").InnerText;

                Debug.WriteLine("Name: " + Name);
                Debug.WriteLine("SN: " + SN);
                Debug.WriteLine("SWVer: " + SWVer);
                Debug.WriteLine("HWVer: " + HWVer);

                txtRouterChamberPerformanceFunctionTestName.Text = Name;
                txtRouterChamberPerformanceFunctionTestSerialNumber.Text = SN;
                txtRouterChamberPerformanceFunctionTestSwVersion.Text = SWVer;
                txtRouterChamberPerformanceFunctionTestHwVersion.Text = HWVer;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATE/FunctionTest/Model " + ex);
            }

            // Time Setting
            XmlNode nodeFunctionTestTimeSetting = doc.SelectSingleNode("/RouterATE/FunctionTest/TimeSetting");
            try
            {
                string PingTimeout = nodeFunctionTestTimeSetting.SelectSingleNode("PingTimeout").InnerText;
                string ConditionTimeout = nodeFunctionTestTimeSetting.SelectSingleNode("ConditionTimeout").InnerText;
                string DutRebootTimeout = nodeFunctionTestTimeSetting.SelectSingleNode("DutRebootTimeout").InnerText;

                Debug.WriteLine("PingTimeout: " + PingTimeout);
                Debug.WriteLine("ConditionTimeout: " + ConditionTimeout);
                Debug.WriteLine("DutRebootTimeout: " + DutRebootTimeout);

                nudRouterChamberPerformanceFunctionTestPingTimeout.Value = Convert.ToDecimal(PingTimeout);
                nudRouterChamberPerformanceFunctionTestConditionTimeout.Value = Convert.ToDecimal(ConditionTimeout);
                nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value = Convert.ToDecimal(DutRebootTimeout);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATE/FunctionTest/TimeSetting " + ex);
            }

            // Loop Setting
            XmlNode nodeFunctionTestLoopSetting = doc.SelectSingleNode("/RouterATE/FunctionTest/LoopSetting");
            try
            {
                string ScheduleOnOff = nodeFunctionTestLoopSetting.SelectSingleNode("ScheduleOnOff").InnerText;
                string ScheduleTimes = nodeFunctionTestLoopSetting.SelectSingleNode("ScheduleTimes").InnerText;

                Debug.WriteLine("ScheduleOnOff: " + ScheduleOnOff);
                Debug.WriteLine("ScheduleTimes: " + ScheduleTimes);

                chkRouterChamberPerformanceFunctionTestScheduleOnOff.Checked = (ScheduleOnOff == "On" ? true : false);
                nudRouterChamberPerformanceFunctionTestTimes.Value = Convert.ToDecimal(ScheduleTimes);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATE/FunctionTest/LoopSetting " + ex);
            }

            //Email Section 
            XmlNode nodeConfigurationEmailSetting = doc.SelectSingleNode("/RouterATE/Configuration/EmailSetting");
            try
            {
                string SendReport = nodeConfigurationEmailSetting.SelectSingleNode("SendReport").InnerText;
                string ReportSenderAccount = nodeConfigurationEmailSetting.SelectSingleNode("ReportSenderAccount").InnerText;
                string ReportSenderPassword = nodeConfigurationEmailSetting.SelectSingleNode("ReportSenderPassword").InnerText;
                string ReportSendTo = nodeConfigurationEmailSetting.SelectSingleNode("ReportSendTo").InnerText;

                Debug.WriteLine("SendReport: " + SendReport);
                Debug.WriteLine("ReportSendTo: " + ReportSendTo);
                Debug.WriteLine("ReportSenderAccount: " + ReportSenderAccount);
                Debug.WriteLine("ReportSenderPassword: " + ReportSenderPassword);                
                
                if (SendReport == "Y")
                {
                    chkRouterChamberPerformanceConfigurationSendReport.Checked = true;
                }
                else
                {
                    chkRouterChamberPerformanceConfigurationSendReport.Checked = false;
                }

                txtRouterChamberPerformanceConfigurationReportEmailSenderGmailAccount.Text = ReportSenderAccount;
                txtRouterChamberPerformanceConfigurationReportEmailSenderGmailPassword.Text = ReportSenderPassword;
                txtRouterChamberPerformanceConfigurationReportEmailSendTo.Text = ReportSendTo;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATE/Configuration/EmailSetting " + ex);
            }

            //Line Section 
            XmlNode nodeConfigurationLineSetting = doc.SelectSingleNode("/RouterATE/Configuration/LineSetting");
            try
            {
                string LineMessage = nodeConfigurationLineSetting.SelectSingleNode("LineMessage").InnerText;
                string LineGroupXmlFile = nodeConfigurationLineSetting.SelectSingleNode("LineGroupXmlFile").InnerText;

                Debug.WriteLine("LineMessage: " + LineMessage);
                Debug.WriteLine("LineGroupXmlFile: " + LineGroupXmlFile);

                if (LineMessage == "Y")
                {
                    chkRouterChamberPerformanceConfigurationLineMessage.Checked = true;
                }
                else
                {
                    chkRouterChamberPerformanceConfigurationLineMessage.Checked = false;
                }
                
                txtRouterChamberPerformanceConfigurationLineMessageXmlFileName.Text = LineGroupXmlFile;
                
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATE/Configuration/LineSetting " + ex);
            }

            // Switch Setting 
            XmlNode nodeConfigurationSwitchSetting = doc.SelectSingleNode("/RouterATE/Configuration/SwitchSetting");
            try
            {
                string UseSwitch = nodeConfigurationSwitchSetting.SelectSingleNode("UseSwitch").InnerText;
                string Username = nodeConfigurationSwitchSetting.SelectSingleNode("Username").InnerText;
                string Password = nodeConfigurationSwitchSetting.SelectSingleNode("Password").InnerText;

                Debug.WriteLine("UseSwitch: " + UseSwitch);
                Debug.WriteLine("Username: " + Username);
                Debug.WriteLine("Password: " + Password);

                chkRouterChamberPerformanceConfigurationUseSwitch.Checked = (UseSwitch == "Y" ? true : false);
                txtRouterChamberPerformanceConfigurationSwitchUsername.Text = Username;
                txtRouterChamberPerformanceConfigurationSwitchPassword.Text = Password;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATE/Configuration/SwitchSetting " + ex);
            }

            // Other Setting 
            XmlNode nodeConfigurationOtherSetting = doc.SelectSingleNode("/RouterATE/Configuration/OtherSetting");
            try
            {
                string RemoteServerIp = nodeConfigurationOtherSetting.SelectSingleNode("RemoteServerIp").InnerText;                
                string RemoteServerPort = nodeConfigurationOtherSetting.SelectSingleNode("RemoteServerPort").InnerText;
                string ClientCardIp = nodeConfigurationOtherSetting.SelectSingleNode("ClientCardIp").InnerText;
                string ChariotPath = nodeConfigurationOtherSetting.SelectSingleNode("ChariotPath").InnerText;
                string TestConditionFileName = nodeConfigurationOtherSetting.SelectSingleNode("TestConditionFileName").InnerText;


                Debug.WriteLine("RemoteServerIp: " + RemoteServerIp);
                Debug.WriteLine("RemoteServerPort: " + RemoteServerPort);
                Debug.WriteLine("ClientCardIp: " + ClientCardIp);
                Debug.WriteLine("ChariotPath: " + ChariotPath);
                Debug.WriteLine("TestConditionFileName: " + TestConditionFileName);
                
                mtbRouterChamberPerformanceConfigurationRemoteControlServerIpAddress.Text = RemoteServerIp;
                nudRouterChamberPerformanceConfigurationRemoteControlServerPort.Text = RemoteServerPort;
                mtbRouterChamberPerformanceConfigurationClientCardIpAddress.Text = ClientCardIp;
                txtRouterChamberPerformanceConfigurationChariotFolder.Text = ChariotPath;
                txtRouterChamberPerformanceConfigurationTestConditionExcelFileName.Text = TestConditionFileName;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATE/Configuration/OtherSetting " + ex);
            }

            // Radius Server Section
            XmlNode nodeConfigurationRadiusServerSetting = doc.SelectSingleNode("/RouterATE/Configuration/RadiusServerSetting");
            try
            {
                string RadiusServerIp = nodeConfigurationRadiusServerSetting.SelectSingleNode("RadiusServerIp").InnerText;
                string RadiusServerPort = nodeConfigurationRadiusServerSetting.SelectSingleNode("RadiusServerPort").InnerText;
                

                Debug.WriteLine("RadiusServerIp: " + RadiusServerIp);
                Debug.WriteLine("RadiusServerPort: " + RadiusServerPort);
                

                mtbRouterChamberPerformanceConfigurationRadiusServerIpAddress.Text = RadiusServerIp;
                nudRouterChamberPerformanceConfigurationRadiusServerPort.Text = RadiusServerPort;
                
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATE/Configuration/RadiusServerSetting " + ex);
            }     


            // End of Read Function Test configuration settings

            return true;
        }

        #endregion

        /*============================================================================================*/
        /*====================================== For Test Area ======================================*/
        /*============================================================================================*/
        #region

                private void ForTestRouterChamberPerformanceTest()
                {
                    string str = string.Empty;



        //            //bRouterChamberPerformanceTestFTThreadRunning = true;
        //            //string sFile = "";
        //            //RouterChamberPerformanceTestMainFunction_Sanity(sFile);

        //            for (int i = 0; i < cableSwIntegrationDutDataSets.Length; i++)
        //            {
        //                cableSwIntegrationDutDataSets[i].TestFinish = false;

        //                //cableSwIntegrationDutDataSets[i].snmp.GetMibNextTest(".1.2", ref str );
        //                //string sFileName = @"E:\MibWalk_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm") + ".txt";

        //                //cableSwIntegrationDutDataSets[i].snmp.GetMibWalkPrivate(".1.3.6.1.4.1.46366.4292.77", ref str, sFileName);
        //            }


        //            DateTime Now = DateTime.Now;

        //            if (Now.Minute >= 58)
        //            {
        //                i_SetCheckHourTime = Now.Hour + 1;
        //                i_SetCheckMinuteTime = 00;
        //            }
        //            else
        //            {
        //                i_SetCheckHourTime = Now.Hour;
        //                i_SetCheckMinuteTime = Now.Minute + 1;
        //            }

        //            i_SetCheckSecondTime = 00;


        //            /* Clear Test finished flag */


        //            //if (chkRouterChamberPerformanceConfigurationUseSwitch.Checked)
        //            //{
        //            //    ///* Create COM port 2 for switch control */
        //            //    comPortCableSwIntegrationSwtich = new Comport2();

        //            //    if (comPortCableSwIntegrationSwtich.isOpen() != true)
        //            //    {
        //            //        MessageBox.Show("COM port 2 is not ready!");
        //            //        //System.Windows.Forms.Cursor.Current = Cursors.Default;
        //            //        //btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
        //            //        return;
        //            //    }






        //            //    int x = 0;
        //            //}


        //            /* Create sub folder for saving the report data */
        //            //m_RouterChamberPerformanceTestSubFolder = createRouterChamberPerformanceTestSubFolder(txtRouterChamberPerformanceFunctionTestName.Text);

        //            //string savePath = createRouterChamberPerformanceTestSavePath(m_RouterChamberPerformanceTestSubFolder,  m_LoopRouterChamberPerformanceTest);
        //            //string savePath = string.Empty;

        //            //initialExcelRouterChamberPerformanceTest(savePath);

        //            /* Save end time */
        //            //excelAppRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
        //            //saveExcelRouterChamberPerformanceTest(savePath);


        //            /* Close excel */
        //            //closeExcelRouterChamberPerformanceTest();
        //        }
        //        private void TestSwitch()
        //        {
        //            string str = string.Empty;

        //            if (chkRouterChamberPerformanceConfigurationUseSwitch.Checked)
        //            {
        //                ///* Create COM port 2 for switch control */
        //                comPortCableSwIntegrationSwtich = new Comport2();

        //                if (comPortCableSwIntegrationSwtich.isOpen() != true)
        //                {
        //                    MessageBox.Show("COM port 2 is not ready!");
        //                    //System.Windows.Forms.Cursor.Current = Cursors.Default;
        //                    //btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
        //                    return;
        //                }

        //                switch_sf302 sf302 = new switch_sf302(comPortCableSwIntegrationSwtich);

        //                str = "Login Switch sf302";
        //                SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
        //                sf302.login("cisco", "sqa11063", ref str);
        //                SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

        //                str = "Turn On Port 2";
        //                SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
        //                sf302.SwitchEthernetPort(2, true, ref str);

        //                str = "Loginout Switch sf302";
        //                SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
        //                sf302.logout(ref str);
        //                SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
        //            }
                }

        #endregion

        /*============================================================================================*/
        /*========================================  The End ==========================================*/
        /*============================================================================================*/























//        //CableSwIntegrationDutDataSets[] cableSwIntegrationAllDutDataSets;
//        CableSwIntegrationDutDataSets[] cableSwIntegrationDutDataSets;
//        CableSwIntegrationVeriwave cableSwIntegrationVeriwave;
        
//        /* Declare Veriwave related parameter */
//        bool bCableSwIntegrationThroughputTest = false;  // Record if the test need to test throughput
//        Thread thread_CableSwIntegrationVeriwaveFT;
//        //bool bRouterChamberPerformanceTestFTThreadRunning = false;

//        /* Data Sets Function Test */
//        

//        /* */
//        //Thread theadRouterChamberPerformanceTestStart;

//        //Excel.Workbook excelWorkBookCableSwIntegrationVeriwave;
//        //Excel.Worksheet excelWorkSheetCableSwIntegrationVeriwave;
//        //Excel.Range excelRangeCableSwIntegrationVeriwave;
        
//        /* Time parameter to check the RD ftp server */
//        int i_SetCheckHourTime = 5;
//        int i_SetCheckMinuteTime = 00;
//        int i_SetCheckSecondTime = 00;
//        System.Timers.Timer t_Timer;

//        int i_CableSwIntegrationConditionParamterNum = 30;
//        int i_CableSwIntegrationCurrentIndex = 0;
//        int i_CableSwIntegrationCurrentDateSetsIndex = 0;
//        int i_RouterChamberPerformanceTestLoop = 0;
//        //Thread threadRouterChamberPerformanceTestRunCondition;
//        string[] sa_ReportCableSwIntegration;
//        string s_CableSwIntegrationFinalReportInfo;

//        // 寄出報告
//        string s_EmailSenderID = string.Empty;
//        string s_EmailSenderPassword = string.Empty;
//        string[] s_EmailReceivers = null;
        
//        //CbtTlvBlderControl c_Tlvblder = null;

//        /* Excel object */
//        //Excel.Application excelAppRouterChamberPerformanceTest;
//        //Excel.Workbook[] excelWorkBookRouterChamberPerformanceTest;
//        //Excel.Worksheet[] excelWorkSheetRouterChamberPerformanceTest;
//        //Excel.Range[] excelRangeRouterChamberPerformanceTest;


        
        
//        string[,] sa_TestConditionSets;



        

//        //string[,] saa_testConfigSanityRouterChamberPerformanceTest;
//        //string[,] saa_testConfigFullRouterChamberPerformanceTest;
//        string[,] testConfigRouterChamberPerformanceTest;
//        string[,] testConfigCableSwIntegrationSanityTest;
        
//        //CbtMibAccess[] st_CbtMibDataConfigReadWriteTest;
//        //CbtMibAccess st_CurrentMib;
//        //string[] SnmpReplaceContent;
//        //bool bConditionStatus = false;
//        //int iCmResetMode = 0; // 0:godo_ds mode, 1: /reset mode
//        //int iMtaResetMode = 0; // 0:emta/reset, 1: /reset mode
        
//        /* For 20180103 Test */
//        //string cableConfigBinFile = System.Windows.Forms.Application.StartupPath + "\\importData\\D30_basic_D4U4_BPI_0_James_150105-bpi-off.cfg";
        
//        //const string s_Error_Invalid_TLV = "invalid tlv-11's";
//        //const string s_Error_TLV_11 = "tlv-11 error:";
//        //string s_Error_TLV11_Result = string.Empty;
//        int i_RouterChamberPerformanceTestIndexMaximun = 1000;
//        //List<CbtMibAccess> ls_CurrentMibs = new List<CbtMibAccess>();

//        //string s_ErrorMsgForReport;          

//        //bool b_CfgNotExist = false;
             
              




//        private void DoRouterChamberPerformanceTestStart()
//        {
//            string str = string.Empty;

//            //string RdPath = "192.168.65.201";
//            string RdPath = txtCableSwIntegrationConfigurationRdFtpServerPath.Text;
//            string RdTempPath = System.Windows.Forms.Application.StartupPath + "\\Temp";
//            //string sDirOrig = "/Technicolor/Taipei_BFC5.7.1mp3_RG/CGA2121";
//            string sDirOrig = "/Technicolor/";
//            if (rbtnCableSwIntegrationConfigurationRdFtpMp3Folder.Checked)
//            {
//                sDirOrig += "Taipei_BFC5.7.1mp3_RG/CGA2121";
//            }
//            else
//            {
//                sDirOrig += "CGA2121_BFC5.7.1mp4_RG/CGA2121";
//            }

//            string sDir = sDirOrig;

//            string valueResponsed = string.Empty;
//            string expectedValue = string.Empty;
//            string mibType = string.Empty;
            
//            List<string> list = new List<string>();
//            bool bRdFwDownloadStatus = true;
//            string[] sInTempFile = new string[30];

//            string sDay = DateTime.Now.ToString("yyyy-MM-dd");
//            string sShorDay = DateTime.Now.ToString("yyMMdd");
            
//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();

//            m_RouterChamberPerformanceTestSubFolder = createRouterChamberPerformanceTestSubFolder(txtRouterChamberPerformanceFunctionTestName.Text);
//            //string saveLogPath = m_RouterChamberPerformanceTestSubFolder + "\\" + "PreProcess.txt";
//            string saveLogPath ;

//            for (int i = 0; i < cableSwIntegrationDutDataSets.Length; i++)
//            {
//                i_CableSwIntegrationCurrentDateSetsIndex = i;

//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    bRouterChamberPerformanceTestFTThreadRunning = false;
//                    //MessageBox.Show("Abort test", "Error");
//                    //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                    //threadRouterChamberPerformanceTestFT.Abort();
//                    return;
//                    // Never go here ...
//                }

//                if (cableSwIntegrationDutDataSets[i].TestFinish)
//                {
//                    continue;
//                }

//                /* Create folder for each test or sku */
//                if (cableSwIntegrationDutDataSets[i].TestType.ToLower().IndexOf("throughput") >= 0)
//                { //Reset Device by SNMP                     

//                    bool isExists;
//                    string subPath = m_RouterChamberPerformanceTestSubFolder + "\\Veriwave";                   
//                    cableSwIntegrationDutDataSets[i].ReportSavePath = subPath;

//                    isExists = System.IO.Directory.Exists(subPath);
//                    if (!isExists)
//                    {
//                        str = "Create Veriwave Folder.";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        System.IO.Directory.CreateDirectory(subPath);
//                        str = "Done.";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }                    
//                }
//                else
//                {
//                    /*  Create Sub-folder in terms of sku model */                    

//                    bool isExists;
//                    string subPath = m_RouterChamberPerformanceTestSubFolder + "\\" + cableSwIntegrationDutDataSets[i].SkuModel;
//                    cableSwIntegrationDutDataSets[i].ReportSavePath = subPath;

//                    isExists = System.IO.Directory.Exists(subPath);
//                    if (!isExists)
//                    {
//                        str = "Create SKU Model Folder: " + cableSwIntegrationDutDataSets[i].SkuModel;
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        System.IO.Directory.CreateDirectory(subPath);
//                        str = "Done.";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }                    
//                }

//                //// for temp test
//                //string ExcelFile = txtCableSwIntegrationConfigurationTestConditionExcelFile.Text;
//                //RouterChamberPerformanceTestMainFunction(ExcelFile);
//                //cableSwIntegrationDutDataSets[i].TestFinish = true;                                

//                saveLogPath = cableSwIntegrationDutDataSets[i].ReportSavePath + "\\" + "HeaderLog_" + DateTime.Now.ToString("yyyyMMdd-hhmmss")+ ".txt";
                
//                str = "Current SKU is: " + cableSwIntegrationDutDataSets[i].SkuModel;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = "Check if folder exists in RD server today.";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                /* Check if FW is in RD server */
//                sDir = sDirOrig;
//                //sDir = sDir + "/CGA2121_" + sSku[i] + "/";
//                if (cableSwIntegrationDutDataSets[i].SkuModel.IndexOf("GERMANY") >= 0)
//                {
//                    sDir = sDir + "/CGA2121_GERMANY/";
//                }
//                else
//                {
//                    sDir = sDir + "/CGA2121_" + cableSwIntegrationDutDataSets[i].SkuModel + "/";
//                }
//                sDir = RdPath + sDir;

//                str = "Ftp Path is: ftp://" + sDir;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                CbtFtpClient ftpclient = new CbtFtpClient(sDir, "cbtsqa", "jenkins");

//                try
//                {
//                    /* Check if the folder exists */
//                    list = ftpclient.GetFtpDirList();
//                    bool bCheckPoint = false;

//                    foreach (string s in list)
//                    {                        
//                        if (s == sDay)
//                        {
//                            sDir += sDay + "/";
//                            bCheckPoint = true;                           
//                            break;
//                        }
//                    }

//                    if (!bCheckPoint)
//                    { //找不到今天日期的資料夾
//                        str = "The Date Folder doesn't exist!! ftp://" + sDir + sDay;
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                        File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                        str = String.Format("Sku: {0} 找不到今天的資料夾", cableSwIntegrationDutDataSets[i].SkuModel);

//                        //寄出報告
//                        try
//                        {
//                            SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                        }
//                        catch (Exception ex)
//                        {
//                            str = "Send Mail Failed: " + ex.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        }

//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
//                        continue;
//                    }

//                    str = "The Date Folder Exists: ftp://" + sDir;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    string sDownloadFileName = cableSwIntegrationDutDataSets[i].FwFileName;
//                    sDownloadFileName = sDownloadFileName.Replace("#DATE", sShorDay);
//                    cableSwIntegrationDutDataSets[i].CurrentFwName = sDownloadFileName;

//                    str = "Check if file exists in RD server today:" + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    ftpclient = null;
//                    ftpclient = new CbtFtpClient(sDir, "cbtsqa", "jenkins");

//                    /* Get File list in today's folder */
//                    list.Clear();
//                    list = ftpclient.GetFtpFileList();
//                    string[] sFileName = new string[20];
//                    int iIndex = 0;
                    
//                    //列出目錄中所有檔案
//                    str = "List all the files in Folder: "  + sDir;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    foreach (string s in list)
//                    {
//                        string[] sFiles = s.Split(' ');
//                        foreach (string st in sFiles)
//                        {
//                            if (st.IndexOf("CGA") >= 0)
//                            {
//                                Invoke(new SetTextCallBackT(SetText), new object[] { st, txtRouterChamberPerformanceFunctionTestInformation });
//                                sFileName[iIndex++] = st;
//                            }
//                        }
//                    }
                    
//                    //確認檔案是否存在
//                    str = "Check if the file exists :" + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    bCheckPoint = false;
//                    foreach (string s in sFileName)
//                    {
//                        if(s == sDownloadFileName)
//                        {
//                            bCheckPoint = true;
//                            break;
//                        }                       
//                    }

//                    if (!bCheckPoint)
//                    { //找不到指定的檔案
//                        str = "FW File doesn't exist!! " + sDownloadFileName;
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                        File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                        str = String.Format("Sku: {0} 找不到FW檔案", cableSwIntegrationDutDataSets[i].SkuModel);

//                        //寄出報告
//                        try
//                        {
//                            SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                        }
//                        catch (Exception ex)
//                        {
//                            str = "Send Mail Failed: " + ex.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        }

//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
//                        continue;
//                    }

//                    str = "Start to Download Fw from RD server: " + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    ftpclient = null;
//                    ftpclient = new CbtFtpClient(RdPath, "cbtsqa", "jenkins");
//                    str = "Start to Download File: " + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = "";
//                    ftpclient.DownloadFile(RdTempPath, sDir, sDownloadFileName, ref str);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = "Download File Succeed: " + sDownloadFileName;
//                    sInTempFile[i] = sDownloadFileName;

//                    ftpclient = null;
//                }
//                catch (Exception ex)
//                {
//                    str = "RD FW Doanload Exception: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    saveLogPath = m_RouterChamberPerformanceTestSubFolder + "\\" + "PreProcess_" + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".txt";

//                    File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Sku: {0} Download FW from RD server Failed", cableSwIntegrationDutDataSets[i].SkuModel);

//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                    }
//                    catch (Exception exp)
//                    {
//                        str = "Send Mail Failed: " + exp.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }

//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
//                    continue;
//                }

//                // RD FW 下載完成  
//                str = "RD FW Doanload Finished!!! Start to Upload to NFS Server!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                string NfsPath = txtCableSwIntegrationConfigurationNfsServerPath.Text;
//                //上傳 Firmware 檔案到NFS Server 上面去.

//                str = "Check if Local tftp server exists?";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                if (!Directory.Exists(NfsPath))
//                { //連不到NFS的資料夾,等一個小時之後再試
//                    str = "Local Tftp Server is unreachable!!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    saveLogPath = m_RouterChamberPerformanceTestSubFolder + "\\" + "PreProcess_" + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".txt";

//                    File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("NFS server unreachable");

//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                    }
//                    catch (Exception exp)
//                    {
//                        str = "Send Mail Failed: " + exp.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }

//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
//                    continue;
//                }

//                str = "Server exist. Copy file to local tftp server: ";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                string sFwFile = RdTempPath + "\\" + cableSwIntegrationDutDataSets[i].CurrentFwName;

//                str = "Copy File to Local NFS server: " + sFwFile;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                string sDestFile = NfsPath + "\\" + cableSwIntegrationDutDataSets[i].CurrentFwName;

//                //CopyFileToNfsFoler(sFwFile, NfsPath, true);
//                File.Copy(sFwFile, sDestFile, true);

//                str = "Copy File Succeed.";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //刪除已下載FW
//                //File.Delete(sFwFile);

//                str = "Finished!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = "Start to Write FW File Name to each Device Process!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //將要Download的 FW 檔名寫入Device中
//                int snmpVersion = rbtnCableSwIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnCableSwIntegrationConfigurationSnmpV2.Checked ? 2 : 3;

//                string sIp = cableSwIntegrationDutDataSets[i].IpAddress;

//                str = "Ping Device First: " + sIp;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                 //Ping Device First
//                /* Ping ip first */
//                //if (!QuickPingRouterChamberPerformanceTest(mtbRouterChamberPerformanceFunctionTestCmIpAddress.Text, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//                if (!QuickPingRouterChamberPerformanceTest(sIp, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//                {
//                    str = " Failed!! ";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    saveLogPath = m_RouterChamberPerformanceTestSubFolder + "\\" + "PreProcess_" + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".txt";

//                    File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Sku: {0} Ping Device Failed.", cableSwIntegrationDutDataSets[i].SkuModel);

//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                    }
//                    catch (Exception exp)
//                    {
//                        str = "Send Mail Failed: " + exp.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }

//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
//                    continue;
//                }

//                //Setting Tftp Server IP
//                str = "Setting Tftp Server IP to Device: " + cableSwIntegrationDutDataSets[i].SkuModel + ", " + cableSwIntegrationDutDataSets[i].CurrentFwName;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                str = "Tftp Server IP: " + mtbCableSwIntegrationConfigurationServerIp.Text;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerOid.Text.Trim(), mibType, mtbCableSwIntegrationConfigurationServerIp.Text.Trim(), snmpVersion);
//                Thread.Sleep(1000);
//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //Setting Tftp Server Address
//                str = "Setting Tftp Server Address Type: " + txtCableSwIntegrationConfigurationTftpServerTypeValue.Text;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerTypeOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerTypeOID.Text.Trim(), mibType, txtCableSwIntegrationConfigurationTftpServerTypeValue.Text.Trim(), snmpVersion);
//                Thread.Sleep(1000);
//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerTypeOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = "Write Fw File Name to Device: " + cableSwIntegrationDutDataSets[i].SkuModel + ", " + cableSwIntegrationDutDataSets[i].CurrentFwName;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationImageFileOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationImageFileOid.Text.Trim(), mibType, cableSwIntegrationDutDataSets[i].CurrentFwName, snmpVersion);
//                Thread.Sleep(1000);

//                // 讀回FW Name
//                valueResponsed = "";
//                mibType = "";

//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationImageFileOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(sFwFileNameOid, "OctetString", sFwFiles[i], snmpVersion);
//                //setStatus = Snmp_CableMibReadWriteTest.SetMibSingleValueByVersion(sOid, mibType, sWriteValue, snmpVersion);              

//                str = "Download FW and Reboot By SNMP";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                // Set Admin Status to 1 for downloading firmware
//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), mibType, txtCableSwIntegrationConfigurationAdminStatusValue.Text.Trim(), snmpVersion);
//                Thread.Sleep(1000);
//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                
//                if (cableSwIntegrationDutDataSets[i].TestType.ToLower().IndexOf("throughput") >= 0)
//                { //Reset Device by SNMP 
//                    Thread.Sleep(2000);

//                    str = "Before Reset, Check Device Update (Inprocess) Status: ";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    valueResponsed = "";
//                    mibType = "";
//                    cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = "Wait " + nudCableSwIntegrationConfigurationFwUpdateWaitTime.Value.ToString() + " Seconds for FW Update";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    stopWatch.Stop();
//                    stopWatch.Reset();
//                    stopWatch.Restart();

//                    /* Wait for Fw update finish */
//                    while (true)
//                    {
//                        if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                        {
//                            return;
//                            // Never go here ...
//                        }

//                        if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudCableSwIntegrationConfigurationFwUpdateWaitTime.Value * 1000))
//                            break;

//                        //Thread.Sleep(1000);
//                        Thread.Sleep(1000);
//                        str = String.Format(".");
//                        Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }

//                    str = "Check In-Process Status finished:";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    stopWatch.Stop();
//                    stopWatch.Reset();
//                    stopWatch.Restart();

//                    while (true)
//                    {
//                        bool inProcessStatus = true;

//                        if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                        {
//                            return;
//                            // Never go here ...
//                        }

//                        if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudCableSwIntegrationConfigurationInProcessTimeout.Value * 1000))
//                            break;

//                        str = ".";
//                        Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                        try
//                        {
//                            valueResponsed = "";
//                            mibType = "";
//                            cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//                            str = "Mib Type: " + mibType;
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                            str = String.Format("Read Value: {0}", valueResponsed);
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                            if (str != "3") inProcessStatus = false;
//                        }
//                        catch (Exception ex)
//                        {
//                            str = "Get In-Process Error: " + ex.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        }

//                        if (inProcessStatus) break;
//                        Thread.Sleep(1000);
//                    }

//                    //Wait Time for Device Reset
//                    str = "Wait For Device Reset... " + nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString() + " Seconds.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    stopWatch.Stop();
//                    stopWatch.Reset();
//                    stopWatch.Restart();

//                    while (true)
//                    {
//                        if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                        {
//                            bRouterChamberPerformanceTestFTThreadRunning = false;
//                            return;
//                            // Never go here ...
//                        }

//                        if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value * 1000))
//                        {
//                            break;
//                        }

//                        if(stopWatch.ElapsedMilliseconds % (1000*60) == 0)
//                        {
//                            //str = String.Format("{0} Seconds Passed.", stopWatch.ElapsedMilliseconds % (1000));
//                            //Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        }
//                        else
//                        { 
//                            Thread.Sleep(1000);
//                            str = String.Format(".");
//                            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        }
//                    }

//                    str = "Done." + nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString();

//                    // 讀回目前的 Description 以確認 FW 版本正確
//                    str = String.Format("Get Description to check if FW correct.");
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    valueResponsed = "";
//                    mibType = "";

//                    cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationDescriptionOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    //確認FW版本中有要測試的日期

//                    string sFwFileName = cableSwIntegrationDutDataSets[i].CurrentFwName;
//                    sFwFileName = sFwFileName.Substring(0, sFwFileName.Length - 4);

//                    str = "FW File Name should be: " + sFwFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });     

//                    if (valueResponsed.IndexOf(sFwFileName) < 0)
//                    {
//                        str = " Failed!! ";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                        File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                        str = String.Format("Sku: {0} FW File Name of Dut is wrong.", cableSwIntegrationDutDataSets[i].SkuModel);

//                        //寄出報告
//                        try
//                        {
//                            SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                        }
//                        catch (Exception exp)
//                        {
//                            str = "Send Mail Failed: " + exp.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        }

//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
//                        continue;                       
//                    }

//                    str = " succeed!! ";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = "Save Log to HeaderLog.txt";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = "Done." + nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString();

//                    File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                    CableSwIntegrationVeriwaveMainFunction(i);
//                    cableSwIntegrationDutDataSets[i].TestFinish = true;
//                }
//                else
//                {//Read by Comport
//                    /* Set Comport to device setting comport */
//                    str = "Set Comport to : " + cableSwIntegrationDutDataSets[i].ComportNum;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    comPortRouterChamberPerformanceTest.Close();
//                    comPortRouterChamberPerformanceTest.SetPortName(cableSwIntegrationDutDataSets[i].ComportNum);
//                    comPortRouterChamberPerformanceTest.Open();

//                    if (comPortRouterChamberPerformanceTest.isOpen() != true)
//                    {
//                        str = " Failed!! ";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });                       

//                        File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                        str = String.Format("Sku: {0} Comport: {1} Open Failed.", cableSwIntegrationDutDataSets[i].SkuModel, cableSwIntegrationDutDataSets[i].ComportNum);

//                        //寄出報告
//                        try
//                        {
//                            SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                        }
//                        catch (Exception exp)
//                        {
//                            str = "Send Mail Failed: " + exp.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        }

//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
//                        continue;                       
//                    }

//                    str = "Done.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    stopWatch.Stop();
//                    stopWatch.Reset();
//                    stopWatch.Restart();

//                    /* Wait for Fw update finish */
//                    while (true)
//                    {
//                        if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                        {
//                            return;
//                            // Never go here ...
//                        }

//                        if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudCableSwIntegrationConfigurationFwUpdateWaitTime.Value * 1000))
//                            break;

//                        //if (stopWatch.ElapsedMilliseconds % (1000 * 60) == 0)
//                        //{
//                        //    str = String.Format("====ATE MSG ==== :{0} Seconds Passed.", stopWatch.ElapsedMilliseconds % (1000));
//                        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        //}

//                        //Thread.Sleep(1000);
//                        str = comPortRouterChamberPerformanceTest.ReadLine();

//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }

//                    str = "Check Device Update (Inprocess) Status After FW update: ";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    valueResponsed = "";
//                    mibType = "";
//                    cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    //Wait Time for Device Reset
//                    str = "Wait For Device Reset... " + nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString() + " Seconds.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    stopWatch.Stop();
//                    stopWatch.Reset();
//                    stopWatch.Restart();

//                    while (true)
//                    {
//                        if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                        {
//                            bRouterChamberPerformanceTestFTThreadRunning = false;
//                            return;
//                            // Never go here ...
//                        }

//                        if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value * 1000))
//                        {
//                            break;
//                        }

//                        //if (stopWatch.ElapsedMilliseconds % (1000 * 60) == 0)
//                        //{
//                        //    str = String.Format("====ATE MSG ==== :{0} Seconds Passed.", stopWatch.ElapsedMilliseconds % (1000));
//                        //    Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        //}

//                        str = comPortRouterChamberPerformanceTest.ReadLine();

//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }

//                    str = "Done." + nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString();

//                    if (!CGA2121CheckFwFileName_RouterChamberPerformanceTest(i))
//                    { //Read show cfg failed, or the FW filename is wrong 

//                        str = "Show version failed, or the FW filename is wrong";
//                        Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        File.AppendAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);

//                        //寄出報告
//                        try
//                        {
//                            SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Download FW Process Failed", "Copy Rd FW to NFS server Process Failed!!!", saveLogPath);
//                        }
//                        catch (Exception ex)
//                        {
//                            str = "Send Mail Failed: " + ex.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        }

//                        continue;
//                    }

//                    str = "Save Log to HeaderLog.txt";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
                    
//                    i_CableSwIntegrationCurrentDateSetsIndex = i;
//                    string ExcelFile = txtCableSwIntegrationConfigurationTestConditionExcelFile.Text;

//                    //str = "Read Excel Test Condition!!";
//                    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    //if (!File.Exists(ExcelFile))
//                    //{
//                    //    str = "Excel File doesn't Exist: " + ExcelFile;
//                    //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    //    str = "Read next datagridview test condition.";
//                    //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                        
//                    //    continue;
//                    //}

//                    //RouterChamberPerformanceTestMainFunction(ExcelFile);
//                    RouterChamberPerformanceTestMainFunction_Sanity(ExcelFile);
//                    cableSwIntegrationDutDataSets[i].TestFinish = true;
//                }                
//            }

//            i_RouterChamberPerformanceTestLoop++;                
            
//            bool bAllTestFinished = true;
//            for (int i = 0; i < cableSwIntegrationDutDataSets.Length; i++)
//            {// Check if all the device test finished?
//                if (!cableSwIntegrationDutDataSets[i].TestFinish)
//                    bAllTestFinished = false;
//            }

//            if (bAllTestFinished)
//            {
//                // 重新設定下載FW時間設定為明天早上5:00
//                i_SetCheckHourTime = 5;
//                i_SetCheckMinuteTime = 00;
//                i_SetCheckSecondTime = 00;

//                str = "All Device Test Finished!!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });                

//                str = " ===== Test Completed!! =====";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = "Wait another day to Test and Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = "Waiting for time up....";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                return;
//            }

//            if (DateTime.Now.Hour > 18 || i_RouterChamberPerformanceTestLoop >= 6)
//            {
//                // 重新設定下載FW時間設定為明天早上5:00
//                i_SetCheckHourTime = 5;
//                i_SetCheckMinuteTime = 00;
//                i_SetCheckSecondTime = 00;

//                str = "Some Device Test Failed!!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //str = " ===== Test Completed!! =====";
//                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = "Wait another day to Test and Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = "Waiting for time up....";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            }

//            if (!bAllTestFinished)
//            { //Some Device Test Failed.
//                //連不到NFS的資料夾,等一個小時之後再試
//                str = "Some Device Test Failed, Wait for one hour and try again!!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                if(DateTime.Now.Minute >=58)
//                {
//                    i_SetCheckMinuteTime = 2;
//                    i_SetCheckHourTime = DateTime.Now.Hour + 2;
//                }
//                else
//                {
//                    i_SetCheckHourTime = DateTime.Now.Hour + 1;
//                    i_SetCheckMinuteTime = DateTime.Now.Minute;
//                }

                
//                i_SetCheckSecondTime = DateTime.Now.Second;

//                str = "Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                return;
//            }
            

//            //// 準備開始Device測試, 下次下載FW時間設定為明天早上5:00
//            //i_SetCheckHourTime = 5;
//            //i_SetCheckMinuteTime = 00;
//            //i_SetCheckSecondTime = 00;



//            //// Check RD ftp Server if the FW driver exist
//            //str = "Download FW from RD server to NFS Server.";
//            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            ////if (!CopyRdFwToNfsServerRouterChamberPerformanceTest())
//            ////{
//            ////    str = "Copy Rd FW to NFS server Process Failed!!! Wait next Test Time: " + i_SetCheckHourTime + ":" + i_SetCheckMinuteTime + ":" + i_SetCheckSecondTime; ;
//            ////    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            ////    File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//            ////    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//            ////    //寄出報告
//            ////    try
//            ////    {
//            ////        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Download FW Process Failed", "Copy Rd FW to NFS server Process Failed!!!", saveLogPath);
//            ////    }
//            ////    catch (Exception ex)
//            ////    {
//            ////        str = "Send Mail Failed: " + ex.ToString();
//            ////        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            ////    }

//            ////    str = "Setting Time is: " + i_SetCheckHourTime + ":" + i_SetCheckMinuteTime + ":" + i_SetCheckSecondTime;
//            ////    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            ////    str = "Wait Time ....";
//            ////    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            ////    return;
//            ////    //theadRouterChamberPerformanceTestStart.Abort();
//            ////}

//            ////str = "Download FW Succeed!!";
//            ////Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            ////File.WriteAllText(saveLogPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//            ////Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//            ///* Reset Device */
//            //for (int i = 0; i < cableSwIntegrationDutDataSets.Length; i++)
//            //{
//            //    //string valueResponsed = string.Empty;
//            //    //string expectedValue = string.Empty;
//            //    //string mibType = string.Empty;
//            //    string saveLogSubPath = m_RouterChamberPerformanceTestSubFolder + "\\" + cableSwIntegrationDutDataSets[i].SkuModel + "\\" + "HeaderLog.txt";
//            //    string temp = cableSwIntegrationDutDataSets[i].ReportSavePath + "\\" + "HeaderLog.txt";
                
//            //    if (bRouterChamberPerformanceTestFTThreadRunning == false)
//            //    {
//            //        bRouterChamberPerformanceTestFTThreadRunning = false;
//            //        //MessageBox.Show("Abort test", "Error");
//            //        //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//            //        //threadRouterChamberPerformanceTestFT.Abort();
//            //        return;
//            //        // Never go here ...
//            //    }

//            //    if (cableSwIntegrationDutDataSets[i].TestType.ToLower().IndexOf("throughput") >= 0)
//            //    { //Reset Device by SNMP 
//            //        str = "Create Veriwave Folder.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        bool isExists;
//            //        string subPath = m_RouterChamberPerformanceTestSubFolder + "\\Veriwave";
//            //        saveLogSubPath = subPath + "\\" + "HeaderLog.txt";
//            //        cableSwIntegrationDutDataSets[i].ReportSavePath = subPath;

//            //        isExists = System.IO.Directory.Exists(subPath);
//            //        if (!isExists)
//            //            System.IO.Directory.CreateDirectory(subPath);

//            //        str = "Done.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        str = "Download FW and Reboot By SNMP";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        // Download FW and Reboot By SNMP

//            //        int snmpVersion = rbtnCableSwIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnCableSwIntegrationConfigurationSnmpV2.Checked ? 2 : 3;
//            //        //
//            //        // Set Admin Status to 1 for downloading firmware
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//            //        cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), mibType, txtCableSwIntegrationConfigurationAdminStatusValue.Text.Trim(), snmpVersion);
//            //        Thread.Sleep(1000);
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//            //        str = "Mib Type: " + mibType;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        //cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationResetOid.Text.Trim(), dudCableSwIntegrationConfigurationResetType.Text, txtCableSwIntegrationConfigurationResetValue.Text, snmpVersion);
//            //        Thread.Sleep(2000);

//            //        str = "Before Reset, Check Device Update (Inprocess) Status: ";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //        str = "Mib Type: " + mibType;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        str = "Wait " + nudCableSwIntegrationConfigurationFwUpdateWaitTime.Value.ToString() + " Seconds for FW Update";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        stopWatch.Stop();
//            //        stopWatch.Reset();
//            //        stopWatch.Restart();

//            //        /* Wait for Fw update finish */
//            //        while (true)
//            //        {
//            //            if (bRouterChamberPerformanceTestFTThreadRunning == false)
//            //            {
//            //                return;
//            //                // Never go here ...
//            //            }

//            //            if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudCableSwIntegrationConfigurationFwUpdateWaitTime.Value * 1000))
//            //                break;

//            //            //Thread.Sleep(1000);
//            //            Thread.Sleep(1000);
//            //            str = String.Format(".");
//            //            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //        }

//            //        str = "Check In-Process Status finished:";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        stopWatch.Stop();
//            //        stopWatch.Reset();
//            //        stopWatch.Restart();

//            //        while (true)
//            //        {
//            //            bool inProcessStatus = true;

//            //            if (bRouterChamberPerformanceTestFTThreadRunning == false)
//            //            {
//            //                return;
//            //                // Never go here ...
//            //            }

//            //            if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudCableSwIntegrationConfigurationInProcessTimeout.Value * 1000))
//            //                break;

//            //            str = ".";
//            //            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //            try
//            //            {
//            //                valueResponsed = "";
//            //                mibType = "";
//            //                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //                str = "Mib Type: " + mibType;
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //                str = String.Format("Read Value: {0}", valueResponsed);
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //                if (str != "3") inProcessStatus = false;
//            //            }
//            //            catch (Exception ex)
//            //            {
//            //                str = "Get In-Process Error: " + ex.ToString();
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //            }

//            //            if (inProcessStatus) break;
//            //            Thread.Sleep(1000);
//            //        }

//            //        //Wait Time for Device Reset
//            //        str = "Wait For Device Reset... " + nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString() + " Seconds.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        stopWatch.Stop();
//            //        stopWatch.Reset();
//            //        stopWatch.Restart();

//            //        while (true)
//            //        {
//            //            if (bRouterChamberPerformanceTestFTThreadRunning == false)
//            //            {
//            //                bRouterChamberPerformanceTestFTThreadRunning = false;
//            //                return;
//            //                // Never go here ...
//            //            }

//            //            if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value * 1000))
//            //            {
//            //                break;
//            //            }

//            //            Thread.Sleep(1000);
//            //            str = String.Format(".");
//            //            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //        }

//            //        str = "Done." + nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString();

//            //        str = "Save Log to HeaderLog.txt";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                                       
//            //        str = "Done." + nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString();

//            //        File.WriteAllText(saveLogSubPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//            //        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
//            //    }
//            //    else
//            //    { //Reset by Comport

//            //        /*  Create Sub-folder in terms of sku model */
//            //        str = "Create SKU Model Folder: " + cableSwIntegrationDutDataSets[i].SkuModel;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        bool isExists;
//            //        string subPath = m_RouterChamberPerformanceTestSubFolder + "\\" + cableSwIntegrationDutDataSets[i].SkuModel;
//            //        cableSwIntegrationDutDataSets[i].ReportSavePath = subPath;

//            //        isExists = System.IO.Directory.Exists(subPath);
//            //        if (!isExists)
//            //            System.IO.Directory.CreateDirectory(subPath);

//            //        str = "Done.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        /* Set Comport to device setting comport */
//            //        str = "Set Comport to : " + cableSwIntegrationDutDataSets[i].ComportNum;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        comPortRouterChamberPerformanceTest.Close();
//            //        comPortRouterChamberPerformanceTest.SetPortName(cableSwIntegrationDutDataSets[i].ComportNum);
//            //        comPortRouterChamberPerformanceTest.Open();

//            //        if (comPortRouterChamberPerformanceTest.isOpen() != true)
//            //        {
//            //            //MessageBox.Show("COM port is not ready!");
//            //            //System.Windows.Forms.Cursor.Current = Cursors.Default;
//            //            //btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
//            //            return;
//            //        }

//            //        str = "Done.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        str = "Download FW and Reboot By SNMP";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        // Download FW and Reboot By SNMP
//            //        //int snmpVersion = rbtnCableSwIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnCableSwIntegrationConfigurationSnmpV2.Checked ? 2 : 3;
//            //        //cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationResetOid.Text.Trim(), dudCableSwIntegrationConfigurationResetType.Text, txtCableSwIntegrationConfigurationResetValue.Text, snmpVersion);

//            //        int snmpVersion = rbtnCableSwIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnCableSwIntegrationConfigurationSnmpV2.Checked ? 2 : 3;
//            //        //
//            //        // Set Admin Status to 1 for downloading firmware
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//            //        cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), mibType, txtCableSwIntegrationConfigurationAdminStatusValue.Text.Trim(), snmpVersion);
//            //        Thread.Sleep(1000);
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//            //        str = "Mib Type: " + mibType;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        //cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationResetOid.Text.Trim(), dudCableSwIntegrationConfigurationResetType.Text, txtCableSwIntegrationConfigurationResetValue.Text, snmpVersion);
//            //        Thread.Sleep(2000);

//            //        str = "Before Reset, Check Device Update (Inprocess) Status: ";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //        str = "Mib Type: " + mibType;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                    
//            //        str = "Wait " + nudCableSwIntegrationConfigurationFwUpdateWaitTime.Value.ToString() + " Seconds for FW Update";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        stopWatch.Stop();
//            //        stopWatch.Reset();
//            //        stopWatch.Restart();

//            //        /* Wait for Fw update finish */
//            //        while (true)
//            //        {
//            //            if (bRouterChamberPerformanceTestFTThreadRunning == false)
//            //            {
//            //                return;
//            //                // Never go here ...
//            //            }

//            //            if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudCableSwIntegrationConfigurationFwUpdateWaitTime.Value * 1000))
//            //                break;

//            //            //Thread.Sleep(1000);
//            //            str = comPortRouterChamberPerformanceTest.ReadLine();

//            //            if (str != "")
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //        }

//            //        str = "Check Device Update (Inprocess) Status After FW update: ";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //        str = "Mib Type: " + mibType;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        //try
//            //        //{
//            //        //    valueResponsed = "";
//            //        //    mibType = "";
//            //        //    cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //        //    str = "Mib Type: " + mibType;
//            //        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        //    str = String.Format("Read Value: {0}", valueResponsed);
//            //        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        //    if (str != "3") inProcessStatus = false;
//            //        //}
//            //        //catch (Exception ex)
//            //        //{
//            //        //    str = "Get In-Process Error: " + ex.ToString();
//            //        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //        //}


//            //        //stopWatch.Stop();
//            //        //stopWatch.Reset();
//            //        //stopWatch.Restart();
                    
//            //        //while (true)
//            //        //{
//            //        //    bool inProcessStatus = true;

//            //        //    if (bRouterChamberPerformanceTestFTThreadRunning == false)
//            //        //    {
//            //        //        return;
//            //        //        // Never go here ...
//            //        //    }

//            //        //    if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudCableSwIntegrationConfigurationInProcessTimeout.Value * 1000))
//            //        //        break;

//            //        //    str = ".";
//            //        //    Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                        
//            //        //    try
//            //        //    {
//            //        //        valueResponsed = "";
//            //        //        mibType = "";
//            //        //        cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //        //        str = "Mib Type: " + mibType;
//            //        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        //        if (str != "3") inProcessStatus = false;                           
//            //        //    }
//            //        //    catch (Exception ex)
//            //        //    {
//            //        //        str = "Get In-Process Error: " + ex.ToString();
//            //        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //        //    }

//            //        //    if (inProcessStatus) break;
//            //        //    Thread.Sleep(1000);
//            //        //}
                    
//            //        //Wait Time for Device Reset
//            //        str = "Wait For Device Reset... " + nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString() + " Seconds.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        stopWatch.Stop();
//            //        stopWatch.Reset();
//            //        stopWatch.Restart();

//            //        while (true)
//            //        {
//            //            if (bRouterChamberPerformanceTestFTThreadRunning == false)
//            //            {
//            //                bRouterChamberPerformanceTestFTThreadRunning = false;
//            //                return;
//            //                // Never go here ...
//            //            }

//            //            if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value * 1000))
//            //            {
//            //                break;
//            //            }
                        
//            //            str = comPortRouterChamberPerformanceTest.ReadLine();

//            //            if (str != "")
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //        }

//            //        str = "Done." + nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value.ToString();                    

//            //        str = "Save Log to HeaderLog.txt";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //        File.WriteAllText(saveLogSubPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
//            //        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
//            //    }
//            //}
            
//            //str = "Reset Device Finished. Now Start Function Test.";
//            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //threadRouterChamberPerformanceTestFT = new Thread(new ThreadStart(DoRouterChamberPerformanceFunctionTest));
//            //threadRouterChamberPerformanceTestFT.Name = "";
//            //threadRouterChamberPerformanceTestFT.Start();        
//        }
        
//        /* Main function thread */
//        /* Report power thread function */
//        /* For reference , This function want't be used */
//        private void DoRouterChamberPerformanceFunctionTest()
//        {
//            string str = string.Empty;            

//            int m_totalTimes = 1;  // Default test time
//            m_LoopRouterChamberPerformanceTest = 1; // Reset loop counter
//            if (chkRouterChamberPerformanceFunctionTestScheduleOnOff.Checked)                
//                m_totalTimes = Convert.ToInt32(nudRouterChamberPerformanceFunctionTestTimes.Value);                       

//            /* Start Test Loop */
//            do
//            {
//                int iCount = 1 ;// indicate the test condition number running now. 
//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    bRouterChamberPerformanceTestFTThreadRunning = false;
//                    //MessageBox.Show("Abort test", "Error");
//                    //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                    //threadRouterChamberPerformanceTestFT.Abort();
//                    return;
//                    // Never go here ...
//                }

//                /* Read Test Condition and run each excel file as Main function */
//                for (int conditionRow = 0; conditionRow < 1; conditionRow++)
//                {
//                    if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    {
//                        bRouterChamberPerformanceTestFTThreadRunning = false;
//                        //MessageBox.Show("Abort test", "Error");
//                        //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                        //threadRouterChamberPerformanceTestFT.Abort();
//                        return;
//                        // Never go here ...
//                    }

//                    string ExcelFile = txtCableSwIntegrationConfigurationTestConditionExcelFile.Text;

//                    str = "Read Excel Test Condition!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    if (!File.Exists(ExcelFile))
//                    {
//                        str = "Excel File doesn't Exist: " + ExcelFile;
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        str = "Read next datagridview test condition.";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        iCount++;
//                        continue;
//                    }

//                    str = " ===== Start to run Excel item: " + iCount.ToString() + " =====";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });                   

//                    RouterChamberPerformanceTestMainFunction(ExcelFile);               

//                    iCount++;
//                }

//                m_totalTimes--;
//                m_LoopRouterChamberPerformanceTest++;
//            } while (m_totalTimes != 0);

//            str = " ===== Test Completed!! =====";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            str = "Wait another day to Test and Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            str = "Waiting for time up....";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            return;
//            //bRouterChamberPerformanceTestFTThreadRunning = false;
//            //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));                
//        }

//        /* Main Function for Integration */
//        private bool RouterChamberPerformanceTestMainFunction_Sanity(string sExcelFileName)
//        {
//            string str = string.Empty;
//            int iCount = 0; // Indicate how many excel index row has read.
//            int iIndexCount = 0; // Indicate same index has read.

//            int iIndexCurrentRun = 1;
//            string sExcelFile = string.Empty;

//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();

//            /* Initial Value */
//            m_PositionRouterChamberPerformanceTest = 16;
//            //testConfigRouterChamberPerformanceTest = null;

//            string subFolder = m_RouterChamberPerformanceTestSubFolder + "\\" + cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel;
//            string savePath = createRouterChamberPerformanceTestSaveExcelFile(subFolder, cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel, m_LoopRouterChamberPerformanceTest, sExcelFileName);
//            string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";            
//            //string savePath = createRouterChamberPerformanceTestSaveExcelFile(m_RouterChamberPerformanceTestSubFolder, cableSwIntegrationDutDataSets[i].SkuModel, m_LoopRouterChamberPerformanceTest, sExcelFileName);
//            /* initial Excel component */
//            initialExcelRouterChamberPerformanceTest(savePath);

//            try
//            {
//                /* Fill Loop, PowerLevel and Constellation in Excel */
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[9, 3] = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SerialNumber;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 3] = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SwVersion;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[11, 3] = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].HwVersion;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[4, 9] = m_LoopRouterChamberPerformanceTest;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[5, 9] = Path.GetFileName(sExcelFileName);
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[6, 9] = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 9] = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].CurrentFwName;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[8, 9] = "";
                
//            }
//            catch (Exception ex)
//            {
//                Debug.WriteLine(ex);
//            }

//            ///* Read excel content of excel file */
//            //if (!ReadTestConfig_ReadExcelRouterChamberPerformanceTest(sExcelFileName))
//            //{
//            //    str = " Failed!!";
//            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //    //string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//            //    File.WriteAllText(saveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//            //    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//            //    try
//            //    {
//            //        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = "Read File Failed";

//            //        /* Write Console Log File as HyperLink in Excel report */
//            //        xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3];
//            //        xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//            //    }
//            //    catch (Exception ex)
//            //    {
//            //        str = "Write Data to Excel File Exception: " + ex.ToString();
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //    }
//            //    /* Save end time and close the Excel object */
//            //    xls_excelAppRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//            //    xls_excelWorkBookRouterChamberPerformanceTest.Save();
//            //    /* Close excel application when finish one job(power level) */
//            //    closeExcelRouterChamberPerformanceTest();
//            //    Thread.Sleep(3000);
//            //    //寄出報告
//            //    try
//            //    {
//            //        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Read Excel File Failed", "Read Excel File Failed!!!", savePath);
//            //    }
//            //    catch (Exception ex)
//            //    {
//            //        str = "Send Mail Failed: " + ex.ToString();
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //    }
//            //    return false;
//            //}

//            //str = " Succeed!!";
//            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            /* Run Main condition for Excel Data */
//            for (int rowSanity = 0; rowSanity < testConfigCableSwIntegrationSanityTest.GetLength(0); rowSanity++)
//            {                
//                int iIndexRead;
//                string iIndex = string.Empty;
//                string sFunctionName = string.Empty;
//                string sName = string.Empty;
//                bool bHasInfo = false;
//                s_CableSwIntegrationFinalReportInfo = "";

//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    bRouterChamberPerformanceTestFTThreadRunning = false;
//                    //MessageBox.Show("Abort test", "Error");
//                    //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                    //threadRouterChamberPerformanceTestFT.Abort();
//                    return false;
//                    // Never go here ...
//                }

//                if (testConfigCableSwIntegrationSanityTest[rowSanity, 0] == "" || testConfigCableSwIntegrationSanityTest[rowSanity, 0] == null)
//                {                    
//                    continue;
//                }

//                if (!Int32.TryParse(testConfigCableSwIntegrationSanityTest[rowSanity, 0].Trim(), out iIndexRead))
//                {                    
//                    continue;
//                }

//                iIndex = testConfigCableSwIntegrationSanityTest[rowSanity, 0];
//                sFunctionName = testConfigCableSwIntegrationSanityTest[rowSanity, 1];
//                sName = testConfigCableSwIntegrationSanityTest[rowSanity, 2];                
                
//                str = "====Index " + iIndex + " Start=====";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                i_CableSwIntegrationCurrentIndex = iIndexRead;
//                string sLogFile = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_CableSwIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";

//                str = String.Format("Run condition: Function Name: {0}, Name: {1}", sFunctionName, sName);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                
//                try
//                {
//                    /* Fill Loop, PowerLevel and Constellation in Excel */
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = iIndex;
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3] = sFunctionName;
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 4] = sName;                    
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine(ex);
//                }

//                /* Run each Sanity Test Item */
//                int iStartIndex = -1;
//                int iStopIndex = -1;

//                if(!RouterChamberPerformanceTestGetSanityIndex(sFunctionName, ref iStartIndex, ref iStopIndex))
//                {
//                    try
//                    {
//                        /* Fill Loop, PowerLevel and Constellation in Excel */
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5] = "Error";
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 6] = "Function Unknown or Not Support";
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7] = sLogFile;
                        
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

//                        /* Write Console Log File as HyperLink in Excel report */
//                        xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7];
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sLogFile, Type.Missing, "ConsoleLog", "ConsoleLog");            
//                    }
//                    catch (Exception ex)
//                    {
//                        Debug.WriteLine(ex);
//                    }

//                    str = "Function Unknown or Not Support!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
                    
//                    m_PositionRouterChamberPerformanceTest++;
//                    continue;
//                }

//                //switch (sFunctionName)
//                //{
//                //    case "Verify DUT can be reset by GUI":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //        break;
//                //    case "Remote factory reset via SNMP":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_RemoteFactoryResetViaSnmpStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_RemoteFactoryResetViaSnmpStopIndex;
//                //        break;
//                //    case "Upstream and downstream can be locked":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_UpstreamAndDownstreamCanBeLockedStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_UpstreamAndDownstreamCanBeLockedStopIndex;
//                //        break;
//                //    case "MDD in Docsis2.0/3.0 mode MUST be work":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_MddInDocsis20_30ModeMustBeWorkStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_MddInDocsis20_30ModeMustBeWorkStopIndex;
//                //        break;
//                //    case "tchCmAPBpi2CertStatus":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_tchCmAPBpi2CertStatusTestStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_tchCmAPBpi2CertStatusTestStopIndex;
//                //        break;
//                //    case "tchVendorDefaultDSfreq":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_tchVendorDefaultDSfreqTestStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_tchVendorDefaultDSfreqTestStopIndex;
//                //        break;
//                //    case "tchCmForceDualscan":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_tchCmForceDualscanTestStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_tchCmForceDualscanTestStopIndex;
//                //        break;
//                //    case "SNR & DS/US Power":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_SnrAndDsUsPowerStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_SnrAndDsUsPowerStopIndex;
//                //        break;
//                //    case "CM SNMP Agent":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_CmSnmpAgentStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_CmSnmpAgentStopIndex;
//                //        break;
//                //    case "MTA SNMP Agent":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_MtaSnmpAgentStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_MtaSnmpAgentStopIndex;
//                //        break;
//                //    case "SSID and Password generate rule":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_SsidAndPasswordGenerateRuleStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_SsidAndPasswordGenerateRuleStopIndex;
//                //        break;
//                //    case "5 GHz and 2.4 GHz Each band can be selected":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_5g24gEachBandCanBeSelectedStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_5g24gEachBandCanBeSelectedStopIndex;
//                //        break;
//                //    //case "":
//                //    //    iStartIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;

//                //    //    case "":
//                //    //    iStartIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;
//                //    //    case "":
//                //    //    iStartIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;
//                //    //    case "":
//                //    //    iStartIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;
//                //    //    case "":
//                //    //    iStartIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;
//                //    //    case "":
//                //    //    iStartIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_CableSwIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;
//                //    case "For Test":
//                //        iStartIndex = i_CableSwIntegrationSanityFunction_ForTestStartIndex;
//                //        iStopIndex = i_CableSwIntegrationSanityFunction_ForTestStopIndex;
//                //        break;
                    
//                //    default:                        
//                //        break;                        
//                //}

//                //str = String.Format("Sub-Index: Start:{0}, Stop: {1}", iStartIndex.ToString(), iStopIndex.ToString());
//                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //if(iStartIndex == -1 || iStopIndex == -1)
//                //{                   
//                //    try
//                //    {
//                //        /* Fill Loop, PowerLevel and Constellation in Excel */
//                //        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5] = "Error";
//                //        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 6] = "Function Unknown or Not Support";
//                //        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7] = sLogFile;
                        
//                //        xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

//                //        /* Write Console Log File as HyperLink in Excel report */
//                //        xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7];
//                //        xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sLogFile, Type.Missing, "ConsoleLog", "ConsoleLog");            
//                //    }
//                //    catch (Exception ex)
//                //    {
//                //        Debug.WriteLine(ex);
//                //    }

//                //    str = "Function Unknown or Not Support!!";
//                //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //    File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                //    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });
                    
//                //    m_PositionRouterChamberPerformanceTest++;
//                //    continue;
//                //}

//                s_CableSwIntegrationFinalReportInfo = "";
//                bHasInfo = false;

//                /* Run Condition IndexStart to IndexStop */
//                for (int index = iStartIndex; index <= iStopIndex; index++)
//                {
//                    iIndexCount = 0;
//                    sa_TestConditionSets = null;
//                    sa_ReportCableSwIntegration = null;
//                    GC.Collect();

//                    if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    {
//                        bRouterChamberPerformanceTestFTThreadRunning = false;                        
//                        return false;
//                        // Never go here ...
//                    }

//                    str = "== Run sub-Index " + index.ToString() + " Start==";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    for (int row = 0; row < testConfigRouterChamberPerformanceTest.GetLength(0); row++)
//                    {                        
                        
                        
//                        //i_CableSwIntegrationCurrentIndex = index;
//                        sa_ReportCableSwIntegration = new string[i_CableSwIntegrationConditionParamterNum + 3];

//                        if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                        {
//                            bRouterChamberPerformanceTestFTThreadRunning = false;
//                            return false;
//                            // Never go here ...
//                        }

//                        if (testConfigRouterChamberPerformanceTest[row, 0] == "" || testConfigRouterChamberPerformanceTest[row, 0] == null)
//                        {
//                            continue;
//                        }

//                        if (!Int32.TryParse(testConfigRouterChamberPerformanceTest[row, 0].Trim(), out iIndexRead))
//                        {
//                            continue;
//                        }

//                        if (iIndexRead == index)
//                        {
//                            if (sa_TestConditionSets == null)
//                            {
//                                //sa_TestConditionSets = new string[cableSwIntegrationDutDataSets.Length, i_CableSwIntegrationConditionParamterNum];
//                                sa_TestConditionSets = new string[30, i_CableSwIntegrationConditionParamterNum];
//                            }

//                            for (int j = 0; j < testConfigRouterChamberPerformanceTest.GetLength(1); j++)
//                            {
//                                sa_TestConditionSets[iIndexCount, j] = testConfigRouterChamberPerformanceTest[row, j];
//                            }

//                            iIndexCount++;
//                            iCount++;
//                        }
//                    } //End of for- loop: Scan every row of testConfigRouterChamberPerformanceTest

//                    if (sa_TestConditionSets == null || sa_TestConditionSets[0, 1] == null)
//                    { // 找不到有此Index的欄位, 直接判定 Test Failed.
//                        try
//                        {
//                            /* Fill Loop, PowerLevel and Constellation in Excel */
//                            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5] = "Fail";
//                            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 6] = "Sub-Index:" +index.ToString() + "Not Existt";
//                            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7] = sLogFile;

//                            xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

//                            /* Write Console Log File as HyperLink in Excel report */
//                            xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7];
//                            xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sLogFile, Type.Missing, "ConsoleLog", "ConsoleLog");
//                        }
//                        catch (Exception ex)
//                        {
//                            Debug.WriteLine(ex);
//                        }

//                        str = "Sub-Index:" +index.ToString() + "Not Existt";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                        File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                        m_PositionRouterChamberPerformanceTest++;
//                        continue;
//                    }
                    
//                    str = "Current Device Index: " + cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].index.ToString() + ", SKU: " + cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = "Run Sub Function : " + sa_TestConditionSets[0, 1];
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    //string[] sReport = new string[i_CableSwIntegrationConditionParamterNum + 3];

//                    /* Run Sub function */
//                    sa_ReportCableSwIntegration[0] = sa_TestConditionSets[0, 0];
//                    sa_ReportCableSwIntegration[1] = sa_TestConditionSets[0, 1];
//                    sa_ReportCableSwIntegration[2] = sa_TestConditionSets[0, 2];

//                    bool bMultiReport = false;

//                    switch (sa_TestConditionSets[0, 1])
//                    {
//                        case "SNMP":
//                            CableSwIntegrationSnmpFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "CMTS":
//                            CableSwIntegrationCmtsFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "DUTCONSOLE":
//                            CableSwIntegrationDutFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "WEB":
//                            CableSwIntegrationGuiFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            bMultiReport = true;
//                            break;
//                        case "CNR_WEB":
//                            CableSwIntegrationGuiFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            bMultiReport = true;
//                            break;
//                        case "SWITCH":
//                            CableSwIntegrationSwtichFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "REMOTE":
//                            CableSwIntegrationRemoteControlFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "ResetDevice":
//                            CableSwIntegrationResetDeviceFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "Wait":
//                            CableSwIntegrationWaitFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "Ftp":
//                            CableSwIntegrationFtpFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "Ping":
//                            CableSwIntegrationPingFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "EditFile":
//                            CableSwIntegrationEditFileFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "CfgConvert":
//                            CableSwIntegrationConvertCfgFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "DownloadFwFile":
//                            CableSwIntegrationDownloadFwFileFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "CheckVer":
//                            CableSwIntegrationCheckVersionFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "UploadFileToServer":
//                            CableSwIntegrationUploadFileToSerrverFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                            break;


                            
                            
                            


                            
//                        //case "Ping":
//                        //    CableSwIntegrationPingFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                        //    break;
//                        //case "SNMP":
//                        //    break;
//                        //case "SNMP":
//                        //    break;
//                        default:
//                            sa_ReportCableSwIntegration[3] = "Error";
//                            sa_ReportCableSwIntegration[4] = "Function Name Not Support!!";
//                            for (int reportIndex = 6; reportIndex < sa_ReportCableSwIntegration.Length; reportIndex++)
//                            {
//                                if (sa_TestConditionSets[0, reportIndex - 3] != null)
//                                    sa_ReportCableSwIntegration[reportIndex] = sa_TestConditionSets[0, reportIndex - 3];
//                            }
//                            break;
//                    }

//                    if (sa_ReportCableSwIntegration[3] == "INFO")
//                    {
//                        bHasInfo = true;
//                    }

//                    if (sa_ReportCableSwIntegration[3] == "Error")
//                    { //任何一步有錯誤, 一律中斷, 直接寫 Report, 跳下一個Sanity測試項
//                        //string sLogFile = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_CableSwIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//                        sa_ReportCableSwIntegration[5] = sLogFile;
//                        s_CableSwIntegrationFinalReportInfo = sa_ReportCableSwIntegration[4];

//                        RouterChamberPerformanceTestReportData_Sanity(sa_ReportCableSwIntegration);

//                        File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                        index = iStopIndex + 1;
//                        //m_PositionRouterChamberPerformanceTest++;
//                        continue;
//                        //return false;
//                    }
//                    else if (sa_ReportCableSwIntegration[3] == "FAIL")
//                    { //任何一步有錯誤, 一律中斷, 直接寫 Report, 跳下一個Sanity測試項
//                        //string sLogFile = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_CableSwIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//                        sa_ReportCableSwIntegration[5] = sLogFile;
//                        //s_CableSwIntegrationFinalReportInfo = sa_ReportCableSwIntegration[4];
//                        s_CableSwIntegrationFinalReportInfo = "  FAIL in : " + sa_TestConditionSets[0, 2];
//                        //if (s_CableSwIntegrationFinalReportInfo == "")
//                        //{
//                        //    s_CableSwIntegrationFinalReportInfo = "  FAIL in : " + sa_TestConditionSets[0, 2];
//                        //}
//                        //else
//                        //{
//                        //    s_CableSwIntegrationFinalReportInfo += sa_ReportCableSwIntegration[4];
//                        //}

//                        RouterChamberPerformanceTestReportData_Sanity(sa_ReportCableSwIntegration);

//                        File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                        index = iStopIndex + 1;
//                        //m_PositionRouterChamberPerformanceTest++;
//                        continue;
//                        //return false;
//                    }
//                    else if (sa_ReportCableSwIntegration[3] == "FAILFILE")
//                    { //任何一步有錯誤, 一律中斷, 直接寫 Report, 跳下一個Sanity測試項
//                        //string sLogFile = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_CableSwIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//                        sa_ReportCableSwIntegration[5] = sLogFile;
//                        //s_CableSwIntegrationFinalReportInfo = sa_ReportCableSwIntegration[4];

//                        s_CableSwIntegrationFinalReportInfo = "  FAIL in : " + sa_TestConditionSets[0, 2] + sa_ReportCableSwIntegration[4];

//                        //if (s_CableSwIntegrationFinalReportInfo == "")
//                        //{
//                        //    s_CableSwIntegrationFinalReportInfo = "  FAIL in : " + sa_TestConditionSets[0, 2];
//                        //}
//                        //else
//                        //{
//                        //    s_CableSwIntegrationFinalReportInfo += sa_ReportCableSwIntegration[4];
//                        //}

//                        RouterChamberPerformanceTestReportData_Sanity(sa_ReportCableSwIntegration);

//                        File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                        index = iStopIndex + 1;
//                        //m_PositionRouterChamberPerformanceTest++;
//                        continue;
//                        //return false;
//                    }
//                    else
//                    {//PASS
//                        sa_ReportCableSwIntegration[5] = sLogFile;

//                        //if (sa_ReportCableSwIntegration[3] == "FILE" || sa_ReportCableSwIntegration[3] == "PASSFILE")
//                        if (sa_ReportCableSwIntegration[3] == "FILE" || sa_ReportCableSwIntegration[3] == "PASSFILE")
//                        {                       
//                            RouterChamberPerformanceTestReportData_Sanity(sa_ReportCableSwIntegration);
//                            sa_ReportCableSwIntegration[4] = "";
//                        }
//                        s_CableSwIntegrationFinalReportInfo += sa_ReportCableSwIntegration[4];

//                        //測試完成, 寫入最後結果
//                        //string sLogFile = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_CableSwIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//                        //sa_ReportCableSwIntegration[5] = sLogFile;
//                        // 若中間有INFO, Test Result 留白
//                        if (bHasInfo) sa_ReportCableSwIntegration[3] = "";
//                        if (sa_ReportCableSwIntegration[3] == "PASSFILE") sa_ReportCableSwIntegration[3] = "PASS";

//                        RouterChamberPerformanceTestReportData_Sanity(sa_ReportCableSwIntegration);

//                        //File.WriteAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                        File.AppendAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                        //m_PositionRouterChamberPerformanceTest++;
//                    }                   

//                } //End of for-loop: Run IndexStart to IndexStop 
//                m_PositionRouterChamberPerformanceTest++;
//            }
           
//            /* Save end time and close the Excel object */
//            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//            saveExcelRouterChamberPerformanceTest(savePath);
//            //xls_excelWorkBookRouterChamberPerformanceTest.Save();
//            /* Close excel application when finish one job(power level) */
//            closeExcelRouterChamberPerformanceTest();
//            Thread.Sleep(3000);

//            try
//            {
//                SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : " + cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel + " Test Report", "Test Completed!!!", savePath);
//            }
//            catch (Exception ex)
//            {
//                str = "Send Mail Failed: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            }

//            File.AppendAllText(saveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//            Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//            return true;


//            ///* Save end time and close the Excel object */
//            //xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//            //saveExcelRouterChamberPerformanceTest(savePath);
//            ////xls_excelWorkBookRouterChamberPerformanceTest.Save();
//            ///* Close excel application when finish one job(power level) */
//            //closeExcelRouterChamberPerformanceTest();
//            //Thread.Sleep(3000);

//            //try
//            //{
//            //    SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : " + cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel + " Test Report", "Test Completed!!!", savePath);
//            //}
//            //catch (Exception ex)
//            //{
//            //    str = "Send Mail Failed: " + ex.ToString();
//            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });                     
//            //}
            
//            //File.WriteAllText(saveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//            //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//            //return true;
//        }

//        //private bool RouterChamberPerformanceTestMainFunction_Sanity2(int iStartIndex, int iStopIndex)
//        //{

//        //}
        
//        /* For reference , This function want't be used */
//        private bool RouterChamberPerformanceTestMainFunction(string sExcelFileName)
//        {
//            string str = string.Empty;
//            int iCount = 0; // Indicate how many excel index row has read.
//            int iIndexCount = 0; // Indicate same index has read.

//            int iIndexCurrentRun = 1;
//            string sExcelFile = string.Empty;

//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();

//            /* Initial Value */
//            m_PositionRouterChamberPerformanceTest = 16;
//            testConfigRouterChamberPerformanceTest = null;

//            string subFolder = m_RouterChamberPerformanceTestSubFolder + "\\" + cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel;
//            string savePath = createRouterChamberPerformanceTestSaveExcelFile(subFolder, cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel, m_LoopRouterChamberPerformanceTest, sExcelFileName);
//            string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//            //string savePath = createRouterChamberPerformanceTestSaveExcelFile(m_RouterChamberPerformanceTestSubFolder, cableSwIntegrationDutDataSets[i].SkuModel, m_LoopRouterChamberPerformanceTest, sExcelFileName);
//            /* initial Excel component */
//            initialExcelRouterChamberPerformanceTest(savePath);

//            try
//            {
//                /* Fill Loop, PowerLevel and Constellation in Excel */
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[9, 3] = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SerialNumber;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[10, 3] = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SwVersion;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[11, 3] = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].HwVersion;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[4, 9] = m_LoopRouterChamberPerformanceTest;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[5, 9] = Path.GetFileName(sExcelFileName);
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[6, 9] = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 9] = "";
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[8, 9] = "";
//            }
//            catch (Exception ex)
//            {
//                Debug.WriteLine(ex);
//            }

//            /* Read excel content of excel file */
//            if (!ReadTestConfig_ReadExcelRouterChamberPerformanceTest(sExcelFileName))
//            {
//                str = " Failed!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                //string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//                File.WriteAllText(saveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                try
//                {
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = "Read File Failed";

//                    /* Write Console Log File as HyperLink in Excel report */
//                    xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3];
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//                }
//                catch (Exception ex)
//                {
//                    str = "Write Data to Excel File Exception: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                }
//                /* Save end time and close the Excel object */
//                xls_excelAppRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                xls_excelWorkBookRouterChamberPerformanceTest.Save();
//                /* Close excel application when finish one job(power level) */
//                closeExcelRouterChamberPerformanceTest();
//                Thread.Sleep(3000);
//                //寄出報告
//                try
//                {
//                    SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Read Excel File Failed", "Read Excel File Failed!!!", savePath);
//                }
//                catch (Exception ex)
//                {
//                    str = "Send Mail Failed: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                }
//                return false;
//            }

//            str = " Succeed!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            /* Run Main condition for Excel Data */
//            for (int index = 1; index <= i_RouterChamberPerformanceTestIndexMaximun; index++)
//            {
//                iIndexCount = 0;
//                sa_TestConditionSets = null;
//                sa_ReportCableSwIntegration = null;
//                GC.Collect();
//                i_CableSwIntegrationCurrentIndex = index;
//                sa_ReportCableSwIntegration = new string[i_CableSwIntegrationConditionParamterNum + 3];


//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    bRouterChamberPerformanceTestFTThreadRunning = false;
//                    //MessageBox.Show("Abort test", "Error");
//                    //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                    //threadRouterChamberPerformanceTestFT.Abort();
//                    return false;
//                    // Never go here ...
//                }

//                /* Check if all the excel content has tested. */
//                int iTotalCount = testConfigRouterChamberPerformanceTest.GetLength(0);
//                if (iCount >= testConfigRouterChamberPerformanceTest.GetLength(0)) break;

//                str = "====Index " + index.ToString() + " Start=====";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                for (int row = 0; row < testConfigRouterChamberPerformanceTest.GetLength(0); row++)
//                {
//                    int iIndexRead;

//                    if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    {
//                        bRouterChamberPerformanceTestFTThreadRunning = false;
//                        //MessageBox.Show("Abort test", "Error");
//                        //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                        //threadRouterChamberPerformanceTestFT.Abort();
//                        return false;
//                        // Never go here ...
//                    }

//                    if (testConfigRouterChamberPerformanceTest[row, 0] == "" || testConfigRouterChamberPerformanceTest[row, 0] == null)
//                    {
//                        if (index == 1)
//                        { //Count only once
//                            iCount++;
//                        }
//                        continue;
//                    }

//                    if (!Int32.TryParse(testConfigRouterChamberPerformanceTest[row, 0].Trim(), out iIndexRead))
//                    {
//                        if (index == 1)
//                        { //Count only once
//                            iCount++;
//                        }
//                        continue;
//                    }

//                    if (iIndexRead == index)
//                    {
//                        if (sa_TestConditionSets == null)
//                        {
//                            //sa_TestConditionSets = new string[cableSwIntegrationDutDataSets.Length, i_CableSwIntegrationConditionParamterNum];
//                            sa_TestConditionSets = new string[30, i_CableSwIntegrationConditionParamterNum];
//                        }

//                        for (int j = 0; j < testConfigRouterChamberPerformanceTest.GetLength(1); j++)
//                        {
//                            sa_TestConditionSets[iIndexCount, j] = testConfigRouterChamberPerformanceTest[row, j];
//                        }

//                        iIndexCount++;
//                        iCount++;
//                    }
//                } //End of excel content reading for-loop

//                if (sa_TestConditionSets == null || sa_TestConditionSets[0, 1] == null) continue;

//                str = "Current Device Index: " + cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].index.ToString() + ", SKU: " + cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = "Run Sub Function : " + sa_TestConditionSets[0, 1];
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                //string[] sReport = new string[i_CableSwIntegrationConditionParamterNum + 3];

//                /* Run Sub function */
//                sa_ReportCableSwIntegration[0] = sa_TestConditionSets[0, 0];
//                sa_ReportCableSwIntegration[1] = sa_TestConditionSets[0, 1];
//                sa_ReportCableSwIntegration[2] = sa_TestConditionSets[0, 2];

//                switch (sa_TestConditionSets[0, 1])
//                {
//                    case "SNMP":
//                        CableSwIntegrationSnmpFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                        break;
//                    case "CMTS":
//                        CableSwIntegrationCmtsFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                        break;
//                    case "DUTCONSOLE":
//                        CableSwIntegrationDutFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                        break;
//                    case "WEB":
//                        //CableSwIntegrationSnmpFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                        break;
//                    case "CNRWEB":
//                        //CableSwIntegrationSnmpFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                        break;                        
//                    case "SWITCH":
//                        CableSwIntegrationSwtichFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                        break;                        
//                    case "REMOTETCP":
//                        //CableSwIntegrationSnmpFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                        break;
//                    //case "SNMP":
//                    //    break;
//                    //case "SNMP":
//                    //    break;
//                    default:
//                        sa_ReportCableSwIntegration[3] = "Error";
//                        sa_ReportCableSwIntegration[4] = "Function Name Not Support!!";
//                        for (int reportIndex = 6; reportIndex < sa_ReportCableSwIntegration.Length; reportIndex++)
//                        {
//                            if (sa_TestConditionSets[0, reportIndex - 3] != null)
//                                sa_ReportCableSwIntegration[reportIndex] = sa_TestConditionSets[0, reportIndex - 3];
//                        }
//                        break;
//                }

//                string sLogFile = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_CableSwIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//                sa_ReportCableSwIntegration[5] = sLogFile;

//                RouterChamberPerformanceTestReportData(sa_ReportCableSwIntegration);

//                File.WriteAllText(sLogFile, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                //if (threadRouterChamberPerformanceTestRunCondition != null)
//                //    threadRouterChamberPerformanceTestRunCondition.Abort();

//                m_PositionRouterChamberPerformanceTest++;
//            }




//            ///* Wait Time to process next mib */
//            //str = " Wait for Test Condition Finished or Timeout(S):" + nudRouterChamberPerformanceFunctionTestConditionTimeout.Value.ToString();
//            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //bConditionStatus = false;

//            //Thread threadRouterChamberPerformanceTestRunCondition = new Thread(DoRouterChamberPerformanceTestRunCondition);
//            //threadRouterChamberPerformanceTestRunCondition.Name = "RouterChamberPerformanceTestRunCondition";
//            //threadRouterChamberPerformanceTestRunCondition.Start();

//            //try
//            //{
//            //    stopWatch.Stop();
//            //    stopWatch.Reset();
//            //    stopWatch.Start();

//            //    while (true)
//            //    {
//            //        if (bRouterChamberPerformanceTestFTThreadRunning == false)
//            //        {
//            //            bRouterChamberPerformanceTestFTThreadRunning = false;
//            //            //MessageBox.Show("Abort test", "Error");
//            //            //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//            //            //threadRouterChamberPerformanceTestFT.Abort();
//            //            return false;
//            //            // Never go here ...
//            //        }

//            //        if (bConditionStatus)
//            //            break;

//            //        if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestConditionTimeout.Value * 1000))
//            //        {
//            //            if (!bConditionStatus)
//            //            {
//            //                str = "Run Condition Timeout!!";
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //                //if (threadRouterChamberPerformanceTestRunCondition != null)
//            //                //    threadRouterChamberPerformanceTestRunCondition.Abort();
//            //            }
//            //            break;
//            //        }
//            //    }
//            //}
//            //catch (Exception ex)
//            //{
//            //    str = "Exception: Run Condition Timeout: " + ex.ToString();
//            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //}

//            //if (!bConditionStatus)
//            //{ // Report Condition Timeout
//            //    for (int reportIndex = 0; reportIndex < 3; reportIndex++)
//            //    {
//            //        sa_ReportCableSwIntegration[reportIndex] = sa_TestConditionSets[0, reportIndex];
//            //    }

//            //    sa_ReportCableSwIntegration[3] = "Error";
//            //    sa_ReportCableSwIntegration[4] = "Run Condition Timeout";

//            //    for (int reportIndex = 6; reportIndex < sa_ReportCableSwIntegration.Length; reportIndex++)
//            //    {
//            //        if (sa_TestConditionSets[0, reportIndex - 3] != null)
//            //            sa_ReportCableSwIntegration[reportIndex] = sa_TestConditionSets[0, reportIndex - 3];
//            //    }

//            //    //RouterChamberPerformanceTestReportData(sReport);
//            //}

//            //string sLogFile = s_CableIntegrationSnmpsaveFileLog = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_CableSwIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//            //sa_ReportCableSwIntegration[5] = sLogFile;

//            //RouterChamberPerformanceTestReportData(sa_ReportCableSwIntegration);

//            //File.WriteAllText(s_CableIntegrationSnmpsaveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//            //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//            ////if (threadRouterChamberPerformanceTestRunCondition != null)
//            ////    threadRouterChamberPerformanceTestRunCondition.Abort();

//            //    m_PositionRouterChamberPerformanceTest++;
//            //}
//            /* Save end time and close the Excel object */
//            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//            saveExcelRouterChamberPerformanceTest(savePath);
//            //xls_excelWorkBookRouterChamberPerformanceTest.Save();
//            /* Close excel application when finish one job(power level) */
//            closeExcelRouterChamberPerformanceTest();
//            Thread.Sleep(3000);

//            try
//            {
//                SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : " + cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].SkuModel + " Test Report", "Test Completed!!!", savePath);
//            }
//            catch (Exception ex)
//            {
//                str = "Send Mail Failed: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            }

//            File.WriteAllText(saveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//            Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//            return true;
//        }

//        /* For reference , This function want't be used */
//        private bool RouterChamberPerformanceTestMainFunctionOrig(string sExcelFileName)
//        {
//            string str = string.Empty;
//            int iCount = 0; // Indicate how many excel index row has read.
//            int iIndexCount = 0; // Indicate same index has read.
            
//            int iIndexCurrentRun = 1;
//            string sExcelFile = string.Empty;

//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();            

//            /* Initial Value */
//            m_PositionRouterChamberPerformanceTest = 16;
//            testConfigRouterChamberPerformanceTest = null;          

//            /* Start to run main function for each Data Sets */
//            for (int i = 0; i < cableSwIntegrationDutDataSets.Length; i++)
//            {
//                i_CableSwIntegrationCurrentDateSetsIndex = i;
//                iCount = 0;
//                m_PositionRouterChamberPerformanceTest = 16;
//                string saveLogSubPath = m_RouterChamberPerformanceTestSubFolder + "\\" + cableSwIntegrationDutDataSets[i].SkuModel + "\\" + "HeaderLog.txt";

//                if (cableSwIntegrationDutDataSets[i].TestType.ToLower().IndexOf("throughput") >= 0)
//                {
//                    CableSwIntegrationVeriwaveMainFunction(i_CableSwIntegrationCurrentDateSetsIndex);
//                    continue;
//                }

//                /* Set Comport to device setting comport */
//                str = "Set Comport to : " + cableSwIntegrationDutDataSets[i].ComportNum;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                comPortRouterChamberPerformanceTest.Close();
//                comPortRouterChamberPerformanceTest.SetPortName(cableSwIntegrationDutDataSets[i].ComportNum);
//                comPortRouterChamberPerformanceTest.Open();

//                if (comPortRouterChamberPerformanceTest.isOpen() != true)
//                {
//                    //MessageBox.Show("COM port is not ready!");
//                    //System.Windows.Forms.Cursor.Current = Cursors.Default;
//                    //btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
//                    str = "Set Comport Failed : " + cableSwIntegrationDutDataSets[i].ComportNum;
//                    File.AppendAllText(saveLogSubPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
                    
//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Download FW Process Failed", "Copy Rd FW to NFS server Process Failed!!!", saveLogSubPath);
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Send Mail Failed: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }

//                    continue;
//                }

//                str = "Done.";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });                

//                //if (!CGA2121CheckFwFileName_RouterChamberPerformanceTest())
//                //{ //Read show cfg failed, or the FW filename is wrong 

//                //    str = "Show version failed, or the FW filename is wrong";
//                //    Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                //    File.AppendAllText(saveLogSubPath, txtRouterChamberPerformanceFunctionTestInformation.Text);
                    
//                //    //寄出報告
//                //    try
//                //    {
//                //        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Download FW Process Failed", "Copy Rd FW to NFS server Process Failed!!!", saveLogSubPath);
//                //    }
//                //    catch (Exception ex)
//                //    {
//                //        str = "Send Mail Failed: " + ex.ToString();
//                //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                //    }

//                //    continue;
//                //}

//                /* Create excel report and conosle log path */
//                /* Format : string PathFile = @"\report\SUBFOLDER\FileName_Report_DATE_(LOOP).xlsx";*/
//                string subFolder = m_RouterChamberPerformanceTestSubFolder + "\\" + cableSwIntegrationDutDataSets[i].SkuModel;
//                string savePath = createRouterChamberPerformanceTestSaveExcelFile(subFolder, cableSwIntegrationDutDataSets[i].SkuModel, m_LoopRouterChamberPerformanceTest, sExcelFileName);
//                //string savePath = createRouterChamberPerformanceTestSaveExcelFile(m_RouterChamberPerformanceTestSubFolder, cableSwIntegrationDutDataSets[i].SkuModel, m_LoopRouterChamberPerformanceTest, sExcelFileName);
//                /* initial Excel component */
//                initialExcelRouterChamberPerformanceTest(savePath);

//                try
//                {
//                    /* Fill Loop, PowerLevel and Constellation in Excel */
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[4, 9] = m_LoopRouterChamberPerformanceTest;
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[5, 9] = Path.GetFileName(sExcelFileName);
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[6, 9] = cableSwIntegrationDutDataSets[i].SkuModel;
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 9] = "";
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[8, 9] = "";
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine(ex);
//                }

//                /* Ping Device ip first */
//                str = "Ping Device IP Address: " + cableSwIntegrationDutDataSets[i].IpAddress;
//                Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                if (!QuickPingRouterChamberPerformanceTest(cableSwIntegrationDutDataSets[i].IpAddress, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//                {
//                    str = " Failed!! Run next test condition...";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//                    File.WriteAllText(saveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                    try
//                    {
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = "Ping Failed";

//                        /* Write Console Log File as HyperLink in Excel report */
//                        xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3];
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Write Data to Excel File Exception: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }

//                    /* Save end time and close the Excel object */
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                    xls_excelWorkBookRouterChamberPerformanceTest.Save();
//                    /* Close excel application when finish one job(power level) */
//                    closeExcelRouterChamberPerformanceTest();
//                    Thread.Sleep(3000);

//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Ping Device Failed", "Ping Device Failed!!!", savePath);
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Send Mail Failed: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }

//                    //return false;
//                    continue;
//                }

//                str = " Succeed.";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = "Read Excel Content: " + sExcelFileName;
//                Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                /* Read excel content of excel file */
//                if (!ReadTestConfig_ReadExcelRouterChamberPerformanceTest(sExcelFileName))
//                {
//                    str = " Failed!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//                    File.WriteAllText(saveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                    try
//                    {
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = "Read File Failed";

//                        /* Write Console Log File as HyperLink in Excel report */
//                        xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3];
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Write Data to Excel File Exception: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }
//                    /* Save end time and close the Excel object */
//                    xls_excelAppRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                    xls_excelWorkBookRouterChamberPerformanceTest.Save();
//                    /* Close excel application when finish one job(power level) */
//                    closeExcelRouterChamberPerformanceTest();
//                    Thread.Sleep(3000);

//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Read Excel File Failed", "Read Excel File Failed!!!", savePath);
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Send Mail Failed: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }
//                    continue;
//                    //return false;
//                }

//                str = " Succeed!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                ///* Open Comport of Datasets */
//                //comPortRouterChamberPerformanceTest.ResetPort(cableSwIntegrationDutDataSets[i].ComportNum);
//                //if (comPortRouterChamberPerformanceTest.isOpen() != true)
//                //{
//                //    //MessageBox.Show("COM port is not ready!");
//                //    //System.Windows.Forms.Cursor.Current = Cursors.Default;
//                //    //btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
//                //    //return;
//                //}

//                /* Run Main condition for Excel Data */
//                for (int index = 1; index <= i_RouterChamberPerformanceTestIndexMaximun; index++)
//                {
//                    iIndexCount = 0;
//                    sa_TestConditionSets = null;
//                    i_CableSwIntegrationCurrentIndex = index;                    

//                    if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    {
//                        bRouterChamberPerformanceTestFTThreadRunning = false;
//                        //MessageBox.Show("Abort test", "Error");
//                        //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                        //threadRouterChamberPerformanceTestFT.Abort();
//                        return false;
//                        // Never go here ...
//                    }

//                    /* Check if all the excel content has tested. */
//                    int iTotalCount = testConfigRouterChamberPerformanceTest.GetLength(0);
//                    if (iCount >= testConfigRouterChamberPerformanceTest.GetLength(0)) break;

//                    str = "====Index " + index.ToString() + " Start=====";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    for (int row = 0; row < testConfigRouterChamberPerformanceTest.GetLength(0); row++)
//                    {                        
//                        int iIndexRead;

//                        if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                        {
//                            bRouterChamberPerformanceTestFTThreadRunning = false;
//                            //MessageBox.Show("Abort test", "Error");
//                            //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                            //threadRouterChamberPerformanceTestFT.Abort();
//                            return false;
//                            // Never go here ...
//                        }

//                        if (testConfigRouterChamberPerformanceTest[row, 0] == "" || testConfigRouterChamberPerformanceTest[row, 0] == null)
//                        {
//                            if (index == 1)
//                            { //Count only once
//                                iCount++;
//                            }
//                            continue;
//                        }

//                        if (!Int32.TryParse(testConfigRouterChamberPerformanceTest[row, 0].Trim(), out iIndexRead))
//                        {
//                            if (index == 1)
//                            { //Count only once
//                                iCount++;
//                            }
//                            continue;
//                        }

//                        if (iIndexRead == index)
//                        {
//                            if (sa_TestConditionSets == null)
//                            {
//                                //sa_TestConditionSets = new string[cableSwIntegrationDutDataSets.Length, i_CableSwIntegrationConditionParamterNum];
//                                sa_TestConditionSets = new string[100, i_CableSwIntegrationConditionParamterNum];
//                            }

//                            for (int j = 0; j < testConfigRouterChamberPerformanceTest.GetLength(1); j++)
//                            {
//                                sa_TestConditionSets[iIndexCount, j] = testConfigRouterChamberPerformanceTest[row, j];
//                            }

//                            iIndexCount++;
//                            iCount++;
//                        }
//                    } //End of excel content reading for-loop

//                    if (sa_TestConditionSets == null  || sa_TestConditionSets[0, 1] == null) continue;

//                    str = "Current Device Index: " + cableSwIntegrationDutDataSets[i].index.ToString() + ", SKU: " + cableSwIntegrationDutDataSets[i].SkuModel;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = "Run Sub Function : " + sa_TestConditionSets[0, 1];
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    string[] sReport = new string[i_CableSwIntegrationConditionParamterNum + 3];
                    
//                    /* Wait Time to process next mib */
//                    str = " Wait for Test Condition Finished or Timeout(S):" + nudRouterChamberPerformanceFunctionTestConditionTimeout.Value.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                    
//                    bConditionStatus = false;
                    
//                    try
//                    {
//                        Thread threadRouterChamberPerformanceTestRunCondition = new Thread(DoRouterChamberPerformanceTestRunCondition);
//                        threadRouterChamberPerformanceTestRunCondition.Name = "RouterChamberPerformanceTestRunCondition";
//                        threadRouterChamberPerformanceTestRunCondition.Start();

//                        stopWatch.Stop();
//                        stopWatch.Reset();
//                        stopWatch.Start();

//                        while (true)
//                        {
//                            if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                            {
//                                bRouterChamberPerformanceTestFTThreadRunning = false;
//                                //MessageBox.Show("Abort test", "Error");
//                                //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                                //threadRouterChamberPerformanceTestFT.Abort();
//                                return false;
//                                // Never go here ...
//                            }

//                            if (bConditionStatus)
//                                break;

//                            if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestConditionTimeout.Value * 1000))
//                            {
//                                if (!bConditionStatus)
//                                {
//                                    str = "Run Condition Timeout!!";
//                                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                                    if (threadRouterChamberPerformanceTestRunCondition != null)
//                                        threadRouterChamberPerformanceTestRunCondition.Abort();
//                                }
//                                break;
//                            }
//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Exception: Run Condition Timeout: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    }

//                    if (!bConditionStatus)
//                    { // Report Condition Timeout
//                        for (int reportIndex = 0; reportIndex < 3; reportIndex++)
//                        {
//                            sReport[reportIndex] = sa_TestConditionSets[0, reportIndex] ;
//                        }

//                        sComment = "Error";
//                        sReport[5] = "Run Condition Timeout";

//                        for (int reportIndex = 6; reportIndex < sReport.Length ; reportIndex++)
//                        {
//                            if(sa_TestConditionSets[0, reportIndex -3 ] !=null)
//                                sReport[reportIndex] = sa_TestConditionSets[0, reportIndex-3];
//                        }

//                        RouterChamberPerformanceTestReportData(sReport);
//                    }                    

//                    m_PositionRouterChamberPerformanceTest++;
//                }
//                /* Save end time and close the Excel object */
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                saveExcelRouterChamberPerformanceTest(savePath);
//                //xls_excelWorkBookRouterChamberPerformanceTest.Save();
//                /* Close excel application when finish one job(power level) */
//                closeExcelRouterChamberPerformanceTest();
//                Thread.Sleep(3000);
                               
//                try
//                {
//                    SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Test Report", "Test Completed!!!", savePath);
//                }
//                catch (Exception ex)
//                {
//                    str = "Send Mail Failed: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                }
//            }
           
//            return true;
//        }       
        
//        private void DoRouterChamberPerformanceTestRunCondition()
//        {
//            //string[] sReport = new string[i_CableSwIntegrationConditionParamterNum + 3];
//            //string sLogFile = s_CableIntegrationSnmpsaveFileLog = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_CableSwIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";

//            sa_ReportCableSwIntegration[0] = sa_TestConditionSets[0, 0];
//            sa_ReportCableSwIntegration[1] = sa_TestConditionSets[0, 1];
//            sa_ReportCableSwIntegration[2] = sa_TestConditionSets[0, 2]; 

//            switch (sa_TestConditionSets[0, 1])
//            {

//                case "SNMP":
//                    CableSwIntegrationSnmpFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);                    
//                    break;
//                case "CMTS":
//                    CableSwIntegrationCmtsFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                    break;
//                case "DUTCONSOLE":
//                    CableSwIntegrationDutFunction(ref sa_ReportCableSwIntegration, i_CableSwIntegrationCurrentIndex, i_CableSwIntegrationCurrentDateSetsIndex);
//                    break;
//                case "WEB":
//                    break;
//                //case "SNMP":
//                //    break;
//                //case "SNMP":
//                //    break;
//                default:
//                    sa_ReportCableSwIntegration[3] = "Error";
//                    sa_ReportCableSwIntegration[4] = "Function Name Not Support!!";
//                    break;
//            }
            
//            //sReport[5] = sLogFile;

//            //RouterChamberPerformanceTestReportData(sReport);

//            //File.WriteAllText(s_CableIntegrationSnmpsaveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//            //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//            bConditionStatus = true;

//        }
    
//        //private bool CmReChanProcessRouterChamberPerformanceTest()
//        //{            
//        //    string str = string.Empty;
//        //    //string sGoto_dsCheck = "Moving to Downstream Frequency 333000000 Hz";
//        //    string sGoto_dsCheck = "moving to downstream frequency " + nudRouterChamberPerformanceFunctionTestGotoChannel.Value.ToString() + "000000 hz";
//        //    string sCableConfigCfgFile = System.Windows.Forms.Application.StartupPath + "\\importData\\" + txtRouterChamberPerformanceFunctionTestCfgFileName.Text;

//        //    bool bCheckCrash = false;
//        //    bool bCrash = false;
//        //    bool bReChanProcess = false;
//        //    bool bCheckConfigFile = false;
//        //    bool bRightConfigFile = false;
//        //    string strCompare = string.Empty;
//        //    string sCheckConfigFile = string.Empty;
//        //    string NfsPath = txtCableSwIntegrationConfigurationNfsLocation.Text;
//        //    b_Crash = false;

//        //    int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds

//        //    /* Go to root Directory */
//        //    if (!CGA2121BackToRootDirectory_RouterChamberPerformanceTest("cm"))
//        //    {
//        //        return false;
//        //    }

//        //    // /* Re-Chan Device */
//        //    Stopwatch stopWatch = new Stopwatch();
//        //    stopWatch.Start();

//        //    if (iCmResetMode == 0)
//        //    {  //use dogo_ds freq command                

//        //        //comPortRouterChamberPerformanceTest.DiscardBuffer();            
//        //        //Thread.Sleep(1000);

//        //        //comPortRouterChamberPerformanceTest.WriteLine(" cd /");
//        //        //comPortRouterChamberPerformanceTest.WriteLine("");
//        //        //str = comPortRouterChamberPerformanceTest.ReadLine();
//        //        //if(str != "")
//        //        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //        //stopWatch.Stop();
//        //        //stopWatch.Reset();
//        //        //stopWatch.Restart();

//        //        //while (true)
//        //        //{
//        //        //    if (bRouterChamberPerformanceTestFTThreadRunning == false)
//        //        //        return false;

//        //        //    if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//        //        //    {                    
//        //        //        str = "Change to directory: /CM Timeout." ;
//        //        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation }); 
//        //        //       return false;
//        //        //    }

//        //        //    if (str.ToLower().IndexOf("cm>") != -1) 
//        //        //    {
//        //        //        break;
//        //        //    }
//        //        //    else if (str.ToLower().IndexOf("rg>") != -1)
//        //        //    {
//        //        //        comPortRouterChamberPerformanceTest.WriteLine("/Console/sw");
//        //        //        Thread.Sleep(3000);
//        //        //        comPortRouterChamberPerformanceTest.WriteLine("");
//        //        //        str = comPortRouterChamberPerformanceTest.ReadLine();
//        //        //        if(str != "")
//        //        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //        //    }

//        //        //    str = comPortRouterChamberPerformanceTest.ReadLine();
//        //        //    if(str != "")
//        //        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });               
//        //        //}

//        //        str = "Run goto_ds " + nudRouterChamberPerformanceFunctionTestGotoChannel.Value.ToString();
//        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //        str = "";

//        //        stopWatch.Stop();
//        //        stopWatch.Reset();
//        //        stopWatch.Restart();

//        //        long lTimeTemp = stopWatch.ElapsedMilliseconds + 30 * 1000;
//        //        comPortRouterChamberPerformanceTest.WriteLine("/docsis_ctl/goto_ds " + nudRouterChamberPerformanceFunctionTestGotoChannel.Value.ToString());
//        //        str = comPortRouterChamberPerformanceTest.ReadLine();
//        //        if (str != "")
//        //        {
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //        }

//        //        while (true)
//        //        {
//        //            if (bRouterChamberPerformanceTestFTThreadRunning == false)
//        //            {
//        //                bRouterChamberPerformanceTestFTThreadRunning = false;
//        //                MessageBox.Show("Abort test", "Error");
//        //                this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //                threadRouterChamberPerformanceTestFT.Abort();
//        //                // Never go here ...
//        //            }

//        //            if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestGotoWaitTime.Value) * 1000)
//        //            {
//        //                b_ReChanOK = false;
//        //                return false;
//        //            }

//        //            if (str.ToLower().IndexOf(sGoto_dsCheck) != -1)
//        //            {
//        //                b_ReChanOK = true;
//        //                break;
//        //            }

//        //            Thread.Sleep(100);

//        //            if (stopWatch.ElapsedMilliseconds > lTimeTemp)
//        //            {//等待5秒再寫一次, 一直重複寫會有問題
//        //                lTimeTemp = stopWatch.ElapsedMilliseconds + 30 * 1000;
//        //                comPortRouterChamberPerformanceTest.WriteLine("/docsis_ctl/goto_ds " + nudRouterChamberPerformanceFunctionTestGotoChannel.Value.ToString());
//        //            }
//        //            str = comPortRouterChamberPerformanceTest.ReadLine();
//        //            if (str != "")
//        //            {
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            }

//        //        }

//        //        str = "";
//        //        sCheckConfigFile = "";
//        //        bReChanProcess = false;
//        //        b_TlvError = false;
//        //        b_Crash = false;
//        //        //b_ReChanOK = false;

//        //        stopWatch.Stop();
//        //        stopWatch.Reset();
//        //        stopWatch.Restart();

//        //        /* Wait for device chan */
//        //        while (true)
//        //        {
//        //            if (bRouterChamberPerformanceTestFTThreadRunning == false)
//        //            {
//        //                bRouterChamberPerformanceTestFTThreadRunning = false;
//        //                MessageBox.Show("Abort test", "Error");
//        //                this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //                threadRouterChamberPerformanceTestFT.Abort();
//        //                // Never go here ...
//        //            }

//        //            if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestGotoWaitTime.Value) * 1000)
//        //            {
//        //                break;
//        //            }

//        //            str = comPortRouterChamberPerformanceTest.ReadLine();
//        //            if (str != "")
//        //            {
//        //                strCompare += str;
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //                //if (strCompare.IndexOf(sCableDownloadCfg) != -1 || strCompare.ToLower().IndexOf("starting tftp of configuration file") != -1)
//        //                //{  //check if the device crashed or reboot from this text appeared.
//        //                //    bCheckCrash = true;
//        //                //    bReChanProcess = true;
//        //                //}

//        //                //if (strCompare.ToLower().IndexOf("starting tftp of configuration file") != -1)
//        //                //{
//        //                //    bCheckConfigFile = true;
//        //                //    sCheckConfigFile += str;
//        //                //}

//        //                if (strCompare.ToLower().IndexOf(s_Error_Invalid_TLV) != -1)
//        //                {
//        //                    b_TlvError = true;
//        //                }

//        //                if (strCompare.ToLower().IndexOf("crash") != -1 || strCompare.ToLower().IndexOf("bcm338498") != -1)
//        //                {
//        //                    b_Crash = true;
//        //                    break;
//        //                }
//        //                //if (bCheckConfigFile)
//        //                //{
//        //                //    sCheckConfigFile += str;
//        //                //}
//        //            }

//        //            Thread.Sleep(50);
//        //        } //End of while

//        //        if (b_Crash)
//        //        { // Device crash or reboot , copy basic cfg file into tftp server
//        //            str = "System Crash!!! Copy basic cfg file to tftp server.";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            File.Copy(s_CableConfigCfgBasicCfgFile, sCableConfigCfgFile, true);
//        //            CopyFileToNfsFoler(sCableConfigCfgFile, NfsPath, true);
//        //            str = "Copy File Succeed.";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            return false;
//        //        }

//        //        /* Check if the device process re-chan action */
//        //        //if (!bReChanProcess)
//        //        //{
//        //        //    b_ReChanOK = false;                        
//        //        //}
//        //        //else
//        //        //{
//        //        //    b_ReChanOK = true;
//        //        //}

//        //        if (b_TlvError)
//        //        {
//        //            int iIndexStart = strCompare.ToLower().IndexOf(s_Error_Invalid_TLV);
//        //            if (iIndexStart >= 0)
//        //            {
//        //                string sTemp = strCompare.Substring(iIndexStart, strCompare.Length - iIndexStart);
//        //                int iIndexStop = sTemp.IndexOf(".");
//        //                s_Error_TLV11_Result = sTemp.Substring(0, iIndexStop + 1);
//        //            }
//        //            else
//        //            {
//        //                b_TlvError = false;
//        //            }
//        //        }
//        //    }
//        //    else if (iCmResetMode == 1)
//        //    { //use /reset command 

//        //        // /* Reset Device */
//        //        //Stopwatch stopWatch = new Stopwatch();
//        //        //stopWatch.Start();               

//        //        str = "Reset Device... ";
//        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //        str = "";

//        //        stopWatch.Stop();
//        //        stopWatch.Reset();
//        //        stopWatch.Restart();

//        //        long lTimeTemp = stopWatch.ElapsedMilliseconds + 30 * 1000;
//        //        comPortRouterChamberPerformanceTest.WriteLine("/reset");
//        //        str = comPortRouterChamberPerformanceTest.ReadLine();
//        //        if (str != "")
//        //        {
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //        }

//        //        while (true)
//        //        {
//        //            if (bRouterChamberPerformanceTestFTThreadRunning == false)
//        //            {
//        //                bRouterChamberPerformanceTestFTThreadRunning = false;
//        //                MessageBox.Show("Abort test", "Error");
//        //                this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //                threadRouterChamberPerformanceTestFT.Abort();
//        //                // Never go here ...
//        //            }

//        //            if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestGotoWaitTime.Value) * 1000)
//        //            {
//        //                b_ReChanOK = false;
//        //                return false;
//        //            }

//        //            if (str.ToLower().IndexOf("chip id") != -1)
//        //            {
//        //                b_ReChanOK = true;
//        //                break;
//        //            }

//        //            Thread.Sleep(100);

//        //            if (stopWatch.ElapsedMilliseconds > lTimeTemp)
//        //            {//等待5秒再寫一次, 一直重複寫會有問題
//        //                lTimeTemp = stopWatch.ElapsedMilliseconds + 30 * 1000;
//        //                comPortRouterChamberPerformanceTest.WriteLine("/reset");
//        //            }
//        //            str = comPortRouterChamberPerformanceTest.ReadLine();
//        //            if (str != "")
//        //            {
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            }

//        //        }

//        //        str = "";
//        //        sCheckConfigFile = "";
//        //        bReChanProcess = false;
//        //        b_TlvError = false;
//        //        b_Crash = false;
//        //        //b_ReChanOK = false;

//        //        stopWatch.Stop();
//        //        stopWatch.Reset();
//        //        stopWatch.Restart();

//        //        /* Wait for device chan */
//        //        while (true)
//        //        {
//        //            if (bRouterChamberPerformanceTestFTThreadRunning == false)
//        //            {
//        //                bRouterChamberPerformanceTestFTThreadRunning = false;
//        //                MessageBox.Show("Abort test", "Error");
//        //                this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //                threadRouterChamberPerformanceTestFT.Abort();
//        //                // Never go here ...
//        //            }

//        //            if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestGotoWaitTime.Value) * 1000)
//        //            {
//        //                break;
//        //            }

//        //            str = comPortRouterChamberPerformanceTest.ReadLine();
//        //            if (str != "")
//        //            {
//        //                strCompare += str;
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                        
//        //                if (strCompare.ToLower().IndexOf(s_Error_Invalid_TLV) != -1)
//        //                {
//        //                    b_TlvError = true;
//        //                }

//        //                if (strCompare.ToLower().IndexOf("crash") != -1 || strCompare.ToLower().IndexOf("bcm338498") != -1)
//        //                {
//        //                    b_Crash = true;
//        //                    break;
//        //                }
//        //                //if (bCheckConfigFile)
//        //                //{
//        //                //    sCheckConfigFile += str;
//        //                //}
//        //            }

//        //            Thread.Sleep(50);
//        //        } //End of while

//        //        if (b_Crash)
//        //        { // Device crash or reboot , copy basic cfg file into tftp server
//        //            str = "System Crash!!! Copy basic cfg file to tftp server.";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            File.Copy(s_CableConfigCfgBasicCfgFile, sCableConfigCfgFile, true);
//        //            CopyFileToNfsFoler(sCableConfigCfgFile, NfsPath, true);
//        //            str = "Copy File Succeed.";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            return false;
//        //        }

//        //        /* Check if the device process re-chan action */
//        //        //if (!bReChanProcess)
//        //        //{
//        //        //    b_ReChanOK = false;                        
//        //        //}
//        //        //else
//        //        //{
//        //        //    b_ReChanOK = true;
//        //        //}

//        //        if (b_TlvError)
//        //        {
//        //            int iIndexStart = strCompare.ToLower().IndexOf(s_Error_Invalid_TLV);
//        //            if (iIndexStart >= 0)
//        //            {
//        //                string sTemp = strCompare.Substring(iIndexStart, strCompare.Length - iIndexStart);
//        //                int iIndexStop = sTemp.IndexOf(".");
//        //                s_Error_TLV11_Result = sTemp.Substring(0, iIndexStop + 1);
//        //            }
//        //            else
//        //            {
//        //                b_TlvError = false;
//        //            }
//        //        }
//        //    }

//        //    // Check show cfg
//        //    str = "Check show cfg...";
//        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //    //comPortRouterChamberPerformanceTest.WriteLine(" cd /");
//        //    //comPortRouterChamberPerformanceTest.WriteLine("");
//        //    //str = comPortRouterChamberPerformanceTest.ReadLine();
//        //    //if (str != "")
//        //    //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //    //stopWatch.Stop();
//        //    //stopWatch.Reset();
//        //    //stopWatch.Restart();

//        //    //while (true)
//        //    //{
//        //    //    if (bRouterChamberPerformanceTestFTThreadRunning == false)
//        //    //        return false;

//        //    //    if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//        //    //    {
//        //    //        str = "Change to directory: /CM Failed.";
//        //    //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //    //        return false;
//        //    //    }

//        //    //    if (str.ToLower().IndexOf("cm>") != -1)
//        //    //    {
//        //    //        break;
//        //    //    }
//        //    //    else if (str.ToLower().IndexOf("rg") != -1)
//        //    //    {
//        //    //        comPortRouterChamberPerformanceTest.WriteLine("/Console/sw");
//        //    //        Thread.Sleep(3000);
//        //    //        comPortRouterChamberPerformanceTest.WriteLine("");
//        //    //        str = comPortRouterChamberPerformanceTest.ReadLine();
//        //    //        if (str != "")
//        //    //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //    //    }

//        //    //    str = comPortRouterChamberPerformanceTest.ReadLine();
//        //    //    if (str != "")
//        //    //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //    //}

//        //    if (!CGA2121BackToRootDirectory_RouterChamberPerformanceTest("cm"))
//        //    {
//        //        return false;
//        //    }

//        //    bool bCfgStatus = false;
//        //    s_CfgContect = "";
//        //    str = "";
//        //    comPortRouterChamberPerformanceTest.DiscardBuffer();
//        //    Thread.Sleep(1000);
//        //    comPortRouterChamberPerformanceTest.WriteLine("/Console/cm/show cfg");

//        //    str = comPortRouterChamberPerformanceTest.ReadLine();
//        //    if (str != "")
//        //    {
//        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //        s_CfgContect += str;
//        //    }
      
//        //    stopWatch.Stop();
//        //    stopWatch.Reset();
//        //    stopWatch.Restart();

//        //    while (true)
//        //    {
//        //        if (bRouterChamberPerformanceTestFTThreadRunning == false)
//        //        {
//        //            bRouterChamberPerformanceTestFTThreadRunning = false;
//        //            MessageBox.Show("Abort test", "Error");
//        //            this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //            threadRouterChamberPerformanceTestFT.Abort();
//        //            // Never go here ...
//        //        }

//        //        if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//        //        {
//        //            str = "Show cfg Timeout.";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            return false;
//        //        }

//        //        if (bCfgStatus)
//        //        {
//        //            if (str.ToLower().IndexOf("cm>") != -1)
//        //                //if (sInfo.ToLower().IndexOf("cm/cm_console/cm>") != -1)
//        //                break;
//        //        }
//        //        if (str.ToLower().IndexOf("cm config file name") != -1)
//        //        {                        
//        //            bCfgStatus = true;
//        //        }

//        //        if (str != "" && bCfgStatus)
//        //        {
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            s_CfgContect += str;    
//        //        }

//        //        str = comPortRouterChamberPerformanceTest.ReadLine();
//        //        //Thread.Sleep(100);
//        //    } 

//        //    stopWatch.Stop();
//        //    stopWatch.Reset();

//        //    str = "Check Config File Name...";
//        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //    /* Check if the device loaded cfg file file is equal to CableConfigFile.cfg */
//        //    string sConfigFile = Path.GetFileName(txtRouterChamberPerformanceFunctionTestCfgFileName.Text);
//        //    if (s_CfgContect.IndexOf(sConfigFile) == -1)
//        //    {
//        //        bConfigFileNameFailed = true;
//        //        string sErrorMsg = "Config File Name can't be Found!";
                            
//        //        int iIndexTemp1 = str.ToLower().IndexOf("cm config file name:");
//        //        int iIndexTemp2 = str.ToLower().IndexOf("cm config file contents:");

//        //        if (iIndexTemp1 >= 0 && iIndexTemp2 >= 0)
//        //        {
//        //            sErrorMsg = sCheckConfigFile.Substring(iIndexTemp2, iIndexTemp2 - iIndexTemp1);
//        //        }

//        //        str = "Error Config File Name: " + sErrorMsg;
//        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //        return false;               
//        //    }
//        //    return true;
            
//        //    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //    //RouterChamberPerformanceTestReportData(ls_CurrentMibs, 1, bCrash, str);           
//        //}

//        private bool PingAndReadExcelFileRouterChamberPerformanceTest(string savePath, string sExcelFileName, string ip)
//        {
//            string str = string.Empty;

//            /* Ping ip first */
//            if (s_TestType.ToLower() == "cm")
//            {
//                //str = "Ping CM IP Address: " + mtbRouterChamberPerformanceFunctionTestCmIpAddress.Text;
//            }
//            else
//            {
//                //str = "Ping MTA IP Address: " + mtbRouterChamberPerformanceFunctionTestMtaIpAddress.Text;
//            }
//            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
            
//            //if (!QuickPingRouterChamberPerformanceTest(mtbRouterChamberPerformanceFunctionTestMtaIpAddress.Text, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//            if (!QuickPingRouterChamberPerformanceTest(ip, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//            {
//                str = " Failed!! Run next test condition...";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//                File.WriteAllText(saveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                try
//                {
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = "Ping Failed";

//                    /* Write Console Log File as HyperLink in Excel report */
//                    xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3];
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//                }
//                catch (Exception ex)
//                {
//                    str = "Write Data to Excel File Exception: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                }
//                /* Save end time and close the Excel object */
//                xls_excelAppRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                xls_excelWorkBookRouterChamberPerformanceTest.Save();
//                /* Close excel application when finish one job(power level) */
//                closeExcelRouterChamberPerformanceTest();
//                Thread.Sleep(3000);
//                return false;
//            }

//            str = " Succeed.";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            str = "Read Excel Content: " + sExcelFileName;
//            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            /* Read excel content of excel file */
//            if (!ReadTestConfig_ReadExcelRouterChamberPerformanceTest(sExcelFileName))
//            {
//                str = " Failed!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//                File.WriteAllText(saveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                try
//                {
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = "Read File Failed";

//                    /* Write Console Log File as HyperLink in Excel report */
//                    xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3];
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//                }
//                catch (Exception ex)
//                {
//                    str = "Write Data to Excel File Exception: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                }
//                /* Save end time and close the Excel object */
//                xls_excelAppRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                xls_excelWorkBookRouterChamberPerformanceTest.Save();
//                /* Close excel application when finish one job(power level) */
//                closeExcelRouterChamberPerformanceTest();
//                Thread.Sleep(3000);
//                return false;
//            }

//            str = " Succeed!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            return true;
//        }
        
//        //private bool RouterChamberPerformanceTestMainFunction()
//        //{
//        //    string str = string.Empty;
//        //    string valueResponsed = string.Empty;
//        //    int indexCurrentSnmp = 0; //Read from excel , the index value
//        //    int indexPreviousSnmp = 1; //Record the previous index read from excel           
//        //    int preSnmpIndex = 1;
//        //    int postSnmpIndex = 1;
//        //    //bool bFirstIndex = false;

//        //    //try
//        //    //{
//        //    //    int snmpVersion = rbtnRouterChamberPerformanceFunctionTestSnmpV1.Checked ? 1 : rbtnRouterChamberPerformanceFunctionTestSnmpV2.Checked ? 2 : 3;
//        //    //    Snmp_RouterChamberPerformanceTest = new CBTSnmp();
//        //    //    st_CbtMibDataConfigReadWriteTest = new CbtMibAccess[testConfigRouterChamberPerformanceTest.GetLength(0)];
//        //    //    Snmp_RouterChamberPerformanceTest.Init(mtbRouterChamberPerformanceFunctionTestCmIpAddress.Text, snmpVersion, dudRouterChamberPerformanceFunctionTestCmReadCommunity.Text, dudRouterChamberPerformanceFunctionTestCmWriteCommunity.Text, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestSnmpPort.Value), Convert.ToInt32(nudRouterChamberPerformanceFunctionTestTrapPort.Value));

//        //    //}
//        //    //catch (Exception ex)
//        //    //{
//        //    //    //MessageBox.Show(ex.ToString());
//        //    //    str = "SNMP objext Exception: " + ex.ToString();
//        //    //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //    //}      

//        //    ///
//        //    /// MIB Read Write Config test flow
//        //    ///
//        //    /* Main loop : Get test condition count for testing  */
//        //    for (int row = 0; row < testConfigRouterChamberPerformanceTest.GetLength(0); row++)
//        //    {
//        //        if (bRouterChamberPerformanceTestFTThreadRunning == false)
//        //        {
//        //            bRouterChamberPerformanceTestFTThreadRunning = false;
//        //            MessageBox.Show("Abort test", "Error");
//        //            this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //            threadRouterChamberPerformanceTestFT.Abort();
//        //            // Never go here ...
//        //        }

//        //        str = "====Index " +  indexCurrentSnmp.ToString() +" Start=====";
//        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });


//        //        /* 因為比對的關係, 會造成最後一批無法被處理到, 因此多做這段處理, indexCurrentSnmp+1, 
//        //         * 可以跑else 內的部份, copy檔案及reset dut. 這也是因此row的condition 多加1的原因
//        //         * ,在read excel 中處理, row +1
//        //         * */
//        //        if (row == testConfigRouterChamberPerformanceTest.GetLength(0) - 1)
//        //        {
//        //            //postSnmpIndex = row;
//        //            //string snmpcontent = testConfigRouterChamberPerformanceTest[row, 1];
//        //            //SnmpReplaceContent[row] = snmpcontent;
//        //            //AppendLine2TextFile(cableConfigTxtFile, snmpcontent);
//        //            testConfigRouterChamberPerformanceTest[row, 0] = (indexCurrentSnmp + 1).ToString();
//        //        }

//        //        /* Process the title */
//        //        if (testConfigRouterChamberPerformanceTest[row, 0] == "")
//        //        {
//        //            continue;
//        //        }

//        //        if (testConfigRouterChamberPerformanceTest[row, 0].ToLower().IndexOf("index") >= 0)
//        //        {
//        //            continue;
//        //        }                

//        //        if (!Int32.TryParse(testConfigRouterChamberPerformanceTest[row, 0].Trim(), out indexCurrentSnmp))
//        //        {
//        //            bRouterChamberPerformanceTestFTThreadRunning = false;
//        //            MessageBox.Show("Index Parsing Failed", "Error");
//        //            this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //            threadRouterChamberPerformanceTestFT.Abort();
//        //        }

//        //        //if(!bFirstIndex) 
//        //        //{
//        //        //    indexPreviousSnmp = indexCurrentSnmp;
//        //        //    bFirstIndex = true;
//        //        //}

//        //        if (indexCurrentSnmp == indexPreviousSnmp)
//        //        {
//        //            postSnmpIndex = row;                    
//        //            st_CbtMibDataConfigReadWriteTest[row] = new CbtMibAccess();

//        //            st_CbtMibDataConfigReadWriteTest[row].Name = testConfigRouterChamberPerformanceTest[row, 1];
//        //            st_CbtMibDataConfigReadWriteTest[row].FullName = testConfigRouterChamberPerformanceTest[row, 2];
//        //            st_CbtMibDataConfigReadWriteTest[row].Oid = testConfigRouterChamberPerformanceTest[row, 3];
//        //            if(testConfigRouterChamberPerformanceTest[row, 4] != "")
//        //                st_CbtMibDataConfigReadWriteTest[row].OidPlus = testConfigRouterChamberPerformanceTest[row, 4];
//        //            st_CbtMibDataConfigReadWriteTest[row].Type = testConfigRouterChamberPerformanceTest[row, 5];
//        //            st_CbtMibDataConfigReadWriteTest[row].AccessType = testConfigRouterChamberPerformanceTest[row, 6];
//        //            st_CbtMibDataConfigReadWriteTest[row].Indexes = testConfigRouterChamberPerformanceTest[row, 7];
//        //            //st_CbtMibDataConfigReadWriteTest[row].DefaultValue = testConfigRouterChamberPerformanceTest[row, 8];
//        //            st_CbtMibDataConfigReadWriteTest[row].Module = testConfigRouterChamberPerformanceTest[row, 8];
//        //            st_CbtMibDataConfigReadWriteTest[row].Description = testConfigRouterChamberPerformanceTest[row, 9];
//        //            st_CbtMibDataConfigReadWriteTest[row].WrtieValue = testConfigRouterChamberPerformanceTest[row, 10];
//        //            st_CbtMibDataConfigReadWriteTest[row].ExpectedValue = testConfigRouterChamberPerformanceTest[row, 11];
//        //            int iParseValue;
//        //            if(!Int32.TryParse(testConfigRouterChamberPerformanceTest[row, 12],out iParseValue))
//        //            {
//        //                st_CbtMibDataConfigReadWriteTest[row].WaitTime = 1 ;
//        //            }
//        //            else
//        //            {
//        //                st_CbtMibDataConfigReadWriteTest[row].WaitTime = iParseValue;
//        //            }
                    
//        //            //st_CbtMibDataConfigReadWriteTest[row].ReadValue = testConfigRouterChamberPerformanceTest[row, 1]; ;
//        //            //st_CbtMibDataConfigReadWriteTest[row].Name = testConfigRouterChamberPerformanceTest[row, 1]; ;


//        //            //string sName = testConfigRouterChamberPerformanceTest[row, 1];
//        //            //string sFullName = testConfigRouterChamberPerformanceTest[row, 2];
//        //            //string sOid = testConfigRouterChamberPerformanceTest[row, 3];
//        //            //string SType = testConfigRouterChamberPerformanceTest[row, 4];
//        //            //string sAccessType = testConfigRouterChamberPerformanceTest[row, 5];
//        //            //string sIndexces = testConfigRouterChamberPerformanceTest[row, 6];
//        //            //string sDefaleValue = testConfigRouterChamberPerformanceTest[row, 7];
//        //            //string sMidModule = testConfigRouterChamberPerformanceTest[row, 8];
//        //            //string sDescription = testConfigRouterChamberPerformanceTest[row, 9];                    
//        //            //string sWriteValue = testConfigRouterChamberPerformanceTest[row, 10];
//        //            //string sExpectedValue = testConfigRouterChamberPerformanceTest[row, 11];

//        //            /* Convert AccessType to correct text */
//        //            //switch (sAccessType)
//        //            //{
//        //            //    case "INTERGER":
//        //            //    case "INTERGER32":
//        //            //        sAccessType = "Integer";
//        //            //        break;
//        //            //}


//        //            /* Mib text type : SNMP MIB Object(unknown OID 1.3.6.1.4.1.4413.2.2.2.1.9.1.2.1.0):1.3.6.1.4.1.4413.2.2.2.1.9.1.2.1.0, Integer, 2 */
//        //            //string sOid = st_CbtMibDataConfigReadWriteTest[row].Oid.Substring(1, st_CbtMibDataConfigReadWriteTest[row].Oid.Length - 1);
//        //            string sOid = st_CbtMibDataConfigReadWriteTest[row].Oid + st_CbtMibDataConfigReadWriteTest[row].OidPlus;
//        //            //string sOid = st_CbtMibDataConfigReadWriteTest[row].Oid.Substring(1, st_CbtMibDataConfigReadWriteTest[row].Oid.Length -1) + st_CbtMibDataConfigReadWriteTest[row].OidPlus;
//        //            if (sOid.StartsWith("."))
//        //                sOid = sOid.Substring(1, sOid.Length - 1);
                    
//        //            string sSnmpContent1 = "SNMP MIB Object(unknown OID " + sOid;
//        //            string sSnmpContent2 = "):" + sOid + ", " + st_CbtMibDataConfigReadWriteTest[row].Type + ", " + st_CbtMibDataConfigReadWriteTest[row].WrtieValue;

//        //            string sSnmpContent = sSnmpContent1 + sSnmpContent2;

                    

//        //            //SnmpReplaceContent[row] = sSnmpContent;
//        //            AppendLine2TextFile(cableConfigCfgTxtFile, sSnmpContent);
//        //            continue;
//        //        }
//        //        else
//        //        {
//        //            /* 因為row 走到下一個, 但其實沒處理, 沒有copy 內容, 也沒有把內容放到cfg file 裡. 但因為最後一組是多加的, 
//        //             * 而且在前面已經處理過, 所以不用 row++ */
//        //            if (row != testConfigRouterChamberPerformanceTest.GetLength(0) - 1)
//        //            {
//        //                row--;
//        //            }

//        //            /* Convert txt to bin file */
//        //            if (File.Exists(cableConfigBinFile))
//        //            {
//        //                str = "cfg File Exist! Delete File: " + cableConfigBinFile;
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //                File.Delete(cableConfigBinFile);
//        //            }

//        //            b_CfgNotExist = false;

//        //            RemoveEmptyLinesInFile_Common(cableConfigCfgTxtFile);

//        //            CableConfigFileTxt2Cfg(Path.GetDirectoryName(txtRouterChamberPerformanceFunctionTestCfgConfigConverterFileName.Text), cableConfigCfgTxtFile, cableConfigBinFile, ref str);
//        //            if (str != "")
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //            if (!File.Exists(cableConfigBinFile))
//        //            {
//        //                //bRouterChamberPerformanceTestFTThreadRunning = false;
//        //                //MessageBox.Show("Convert txt to Cfg Failed", "Error");
//        //                //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //                //threadRouterChamberPerformanceTestFT.Abort();

//        //                str = "cfg File Convert fail. File doesn't Exist!!!! Test Next Index======";
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                                               
//        //                b_CfgNotExist = true;

//        //                //RouterChamberPerformanceTestReportData(preSnmpIndex, postSnmpIndex, indexPreviousSnmp, false, str);

//        //                /* Reset CableConfigFile to Cable*/
//        //                File.Copy(cableconfigCfgFileTemplate, cableConfigCfgTxtFile, true);

//        //                preSnmpIndex = postSnmpIndex + 1;
//        //                indexPreviousSnmp++;

//        //                m_PositionRouterChamberPerformanceTest++;

//        //                continue;                        
//        //            }

//        //            /* upload mib.cfg file to \\NFS address\file */
//        //            string NfsPath = txtRouterChamberPerformanceFunctionTestMibNfsLocation.Text;

//        //            str = "Check if tftp server exists?";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //            if (Directory.Exists(NfsPath))
//        //            {
//        //                str = "Server exist. Copy bin file to tftp server: " + cableConfigBinFile;
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //                CopyFileToNfsFoler(cableConfigBinFile, NfsPath, true);
//        //                str = "Copy File Succeed.";
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            }
//        //            else
//        //            {
//        //                str = "Server not found. Test Abort.";
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //                bRouterChamberPerformanceTestFTThreadRunning = false;
//        //                MessageBox.Show("Server is not founded.", "Error");
//        //                this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //                threadRouterChamberPerformanceTestFT.Abort();
//        //            }

//        //            /* Reset Device to get new config file for mib read/write */

//        //            //if (!ChangeCableDirecotory_Common("CM_Console>", comPortRouterChamberPerformanceTest))
//        //            //{
//        //            //    bRouterChamberPerformanceTestFTThreadRunning = false;
//        //            //    MessageBox.Show("Chang Cable Directory Failed", "Error");
//        //            //    this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //            //    threadRouterChamberPerformanceTestFT.Abort();
//        //            //    // Never go here ...
//        //            //}

//        //            /* Re-Chan Device */
//        //            Stopwatch stopWatch = new Stopwatch();
//        //            bool bCheckCrash = false;
//        //            bool bCrash = false;
//        //            bool bReChanProcess = false;
//        //            bool bCheckConfigFile = false;
//        //            bool bRightConfigFile = false;
//        //            string strCompare = string.Empty;
//        //            string sCheckConfigFile = string.Empty;
//        //            //CableChanProcess_Common(nudRouterChamberPerformanceFunctionTestGotoChannel.Value.ToString(), comPortRouterChamberPerformanceTest);
//        //            str = "Change Cable Directory to CM/DOCSIS_CTL.";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //            CableChanDirectoryRouterChamberPerformanceTest(nudRouterChamberPerformanceFunctionTestGotoChannel.Value.ToString(), comPortRouterChamberPerformanceTest);
                    
//        //            comPortRouterChamberPerformanceTest.WriteLine("goto_ds " + nudRouterChamberPerformanceFunctionTestGotoChannel.Value.ToString());

//        //            str = "";
//        //            sCheckConfigFile = "";
//        //            bReChanProcess = false;
//        //            b_TlvError = false;
//        //            stopWatch.Start();

//        //            /* Wait for device chan */
//        //            while (true)
//        //            {
//        //                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//        //                {
//        //                    bRouterChamberPerformanceTestFTThreadRunning = false;
//        //                    MessageBox.Show("Abort test", "Error");
//        //                    this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //                    threadRouterChamberPerformanceTestFT.Abort();
//        //                    // Never go here ...
//        //                }

//        //                if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestGotoWaitTime.Value) *1000)
//        //                {
//        //                    break;
//        //                }

                        
//        //                str = comPortRouterChamberPerformanceTest.ReadLine();
//        //                if(str != "")
//        //                {
//        //                    strCompare += str;
//        //                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //                    if (strCompare.IndexOf(sCableDownloadCfg) != -1 || strCompare.ToLower().IndexOf("starting tftp of configuration file") != -1)
//        //                    {  //check if the device crashed or reboot from this text appeared.
//        //                        bCheckCrash = true;
//        //                        bReChanProcess = true;
//        //                    }

//        //                    if (strCompare.ToLower().IndexOf("starting tftp of configuration file") != -1)
//        //                    {
//        //                        bCheckConfigFile = true;
//        //                        sCheckConfigFile += str;
//        //                    }

//        //                    if (strCompare.ToLower().IndexOf(s_Error_Invalid_TLV) != -1)
//        //                    {
//        //                        b_TlvError = true;
//        //                        //s_Error_TLV11_Result = str;
//        //                    }

//        //                    //if (str.ToLower().IndexOf("config file was read") != -1)
//        //                    //{
//        //                    //    sCheckConfigFile += str;
//        //                    //    bCheckConfigFile = false;
//        //                    //}

//        //                    if (bCheckConfigFile)
//        //                    {
//        //                        sCheckConfigFile += str;
//        //                    }
//        //                }                        

//        //                if (bCheckCrash)
//        //                {
//        //                    if (strCompare.ToLower().IndexOf("crash") != -1 || strCompare.ToLower().IndexOf("bcm338498") != -1)
//        //                    {  // Device crash or reboot , copy basic cfg file into tftp server
//        //                        bCrash = true;
//        //                        File.Copy(s_CableConfigCfgBasicCfgFile, cableConfigBinFile, true);
//        //                        str = "System Crash!!! Copy basic cfg file to tftp server.";
//        //                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //                        CopyFileToNfsFoler(cableConfigBinFile, NfsPath, true);
//        //                        str = "Copy File Succeed.";
//        //                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //                        break;
//        //                    }
//        //                }

//        //                Thread.Sleep(50);
//        //            }                                       

//        //            ///* Check if the device loaded cfg file file is equal to CableConfigFile.cfg */
//        //            //string sConfigFile = Path.GetFileName(cableConfigBinFile);
//        //            //if (sCheckConfigFile.IndexOf(sConfigFile) == -1)
//        //            //{
//        //            //    string sErrorMsg = "Config File Name can't be Found!";
//        //            //    string sTemp = string.Empty;
//        //            //    int iIndexTemp1 = sCheckConfigFile.ToLower().IndexOf("cm cfg file:");
//        //            //    if (iIndexTemp1 >= 0)
//        //            //    {
//        //            //        sTemp = sTemp.Substring(iIndexTemp1, sCheckConfigFile.Length - iIndexTemp1);
//        //            //        int iindexTemp2 = sTemp.IndexOf(".cfg");
//        //            //        if (iindexTemp2 >= 0) sErrorMsg = sTemp.Substring(0, iindexTemp2 + 4);
//        //            //    }                                                

//        //            //    //if (iIndexTemp1 >= 0 && iIndexTemp2 >= 0)
//        //            //    //{
//        //            //    //    sErrorMsg = sCheckConfigFile.Substring(iIndexTemp2, iIndexTemp2 - iIndexTemp1);
//        //            //    //}

//        //            //    bRouterChamberPerformanceTestFTThreadRunning = false;
//        //            //    MessageBox.Show("Config File Name is wrong!! Check the server First!!\r\n" + sErrorMsg, "Error");
//        //            //    this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //            //    threadRouterChamberPerformanceTestFT.Abort();
//        //            //}

//        //            /* Check if the device process re-chan action */
//        //            if (!bReChanProcess)
//        //            {
//        //                b_ReChanOK = false;
//        //                //bRouterChamberPerformanceTestFTThreadRunning = false;
//        //                //MessageBox.Show("Re-Chand Process desn't execute!!", "Error");
//        //                //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //                //threadRouterChamberPerformanceTestFT.Abort();
//        //            }
//        //            else
//        //            {
//        //                b_ReChanOK = true;
//        //            }

//        //            if (b_TlvError)
//        //            {
//        //                int iIndexStart = strCompare.ToLower().IndexOf(s_Error_Invalid_TLV);
//        //                if (iIndexStart >= 0)
//        //                {
//        //                    string sTemp = strCompare.Substring(iIndexStart, strCompare.Length - iIndexStart);
//        //                    int iIndexStop = sTemp.IndexOf(".");
//        //                    s_Error_TLV11_Result = sTemp.Substring(0, iIndexStop + 1);
//        //                }
//        //                else
//        //                {
//        //                    b_TlvError = false;
//        //                }

//        //            }
                    
//        //            if (bCrash)
//        //            {
//        //                //Run Main Function
//        //                str = "";
//        //                //RouterChamberPerformanceTestReportData(preSnmpIndex, postSnmpIndex, indexPreviousSnmp, bCrash, str);
//        //            }
//        //            else
//        //            { // Check show cfg
//        //                str = "Change Cable Directory to CM/CM_Console/CM>.";
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                        
//        //                str = "";
//        //                ChangeToCableConsoleCM(ref str, comPortRouterChamberPerformanceTest);

//        //                str = "Show cfg...";
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //                str = "";
//        //                RunCableShowCfgRouterChamberPerformanceTest(ref str, comPortRouterChamberPerformanceTest);

//        //                /* Check if the device loaded cfg file file is equal to CableConfigFile.cfg */
//        //                string sConfigFile = Path.GetFileName(cableConfigBinFile);
//        //                if (str.IndexOf(sConfigFile) == -1)
//        //                {
//        //                    string sErrorMsg = "Config File Name can't be Found!";
                            
//        //                    int iIndexTemp1 = str.ToLower().IndexOf("cm config file name:");
//        //                    int iIndexTemp2 = str.ToLower().IndexOf("cm config file contents:");

//        //                    if (iIndexTemp1 >= 0 && iIndexTemp2 >= 0)
//        //                    {
//        //                        sErrorMsg = sCheckConfigFile.Substring(iIndexTemp2, iIndexTemp2 - iIndexTemp1);
//        //                    }

//        //                    bRouterChamberPerformanceTestFTThreadRunning = false;
//        //                    MessageBox.Show("Config File Name is wrong!! Check the server First!!\r\n" + sErrorMsg, "Error");
//        //                    this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //                    threadRouterChamberPerformanceTestFT.Abort();
//        //                }

//        //                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //                //RouterChamberPerformanceTestReportData(preSnmpIndex, postSnmpIndex, indexPreviousSnmp, bCrash, str);
//        //            }
                                        

//        //            /* Reset Device to get new config file for mib read/write */
//        //            //str = "Reset Dut:";
//        //            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //            //str = "Read Mib Type:";
//        //            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //            //string mibType = "";
//        //            //valueResponsed = "";
//        //            //bool setStatus = false;
//        //            //Snmp_RouterChamberPerformanceTest.GetMibSingleValueByVersion(txtRouterChamberPerformanceFunctionTestMibResetOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//        //            //if (mibType != "")
//        //            //{
//        //            //    str = "Mib Type: " + mibType;
//        //            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //            //    str = "Set Reset OID.";
//        //            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //            //    setStatus = Snmp_RouterChamberPerformanceTest.SetMibSingleValueByVersion(txtRouterChamberPerformanceFunctionTestMibResetOid.Text.Trim(), mibType, txtRouterChamberPerformanceFunctionTestMibResetValue.Text, snmpVersion);
//        //            //}
//        //            //else
//        //            //{
//        //            //    str = "Set Reset OID.";
//        //            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            //    setStatus = Snmp_RouterChamberPerformanceTest.SetMibSingleValueByVersion(txtRouterChamberPerformanceFunctionTestMibResetOid.Text.Trim(), dudRouterChamberPerformanceFunctionTestMibResetMibType.Text, txtRouterChamberPerformanceFunctionTestMibResetValue.Text, snmpVersion);
//        //            //}
//        //            ////str = "Set Finished, Now Get Value.";

//        //            ///* Wait Device to reboot */
//        //            //Thread.Sleep(3000);
//        //            //bool pingStatus = false;
//        //            //str = String.Format("Wait For Device to Reset for :" + nudRouterChamberPerformanceFunctionTestPingTimeout.Value.ToString() + " Seconds");
//        //            //Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            ////Stopwatch stopWatch = new Stopwatch();
//        //            //stopWatch.Reset();
//        //            //stopWatch.Stop();
//        //            //stopWatch.Start();

//        //            //while (true)
//        //            //{
//        //            //    if (bRouterChamberPerformanceTestFTThreadRunning == false)
//        //            //    {
//        //            //        bRouterChamberPerformanceTestFTThreadRunning = false;
//        //            //        MessageBox.Show("Abort test", "Error");
//        //            //        this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //            //        threadRouterChamberPerformanceTestFT.Abort();
//        //            //        // Never go here ...
//        //            //    }

//        //            //    if (stopWatch.ElapsedMilliseconds > (Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//        //            //    {
//        //            //        break;
//        //            //    }

//        //            //    if (PingClient(mtbRouterChamberPerformanceFunctionTestDeviceAddress.Text, 1000))
//        //            //    {
//        //            //        pingStatus = true;
//        //            //        break;
//        //            //    }

//        //            //    Thread.Sleep(1000);
//        //            //    str = String.Format(".");
//        //            //    Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//        //            //}

//        //            //if (!pingStatus)
//        //            //{
//        //            //    bRouterChamberPerformanceTestFTThreadRunning = false;
//        //            //    str = String.Format("\tPing Failed!!");
//        //            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //            //    MessageBox.Show("Ping Failed, Abort Test!!", "Error");
//        //            //    this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//        //            //    threadRouterChamberPerformanceTestFT.Abort();
//        //            //}

//        //            //str = String.Format("\tPing Succeed!!");
//        //            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//        //            //Run Main Function
//        //            //RouterChamberPerformanceTestReportData(preSnmpIndex, postSnmpIndex, indexPreviousSnmp, bCrash);

//        //            /* Reset CableConfigFile to Cable*/
//        //            File.Copy(cableconfigCfgFileTemplate, cableConfigCfgTxtFile, true);
                    
//        //            preSnmpIndex = postSnmpIndex + 1;
//        //            indexPreviousSnmp++;
//        //        }
                
//        //        m_PositionRouterChamberPerformanceTest++;
//        //    }

//        //    return true;
//        //}

//        /* The function is recorded report data */
//        private void RouterChamberPerformanceTestReportData_Sanity(string[] sReport)
//        {
//            string str = string.Empty;
//            string valueResponsed = string.Empty;
//            string expectedValue = string.Empty;
//            string mibType = string.Empty;
//            int iIndexConsoleLog = 0;
//            string sConsoleLogFile = string.Empty;
//            //s_ErrorMsgForReport = string.Empty;
//            int iConditionTimeout = Convert.ToInt32(nudRouterChamberPerformanceFunctionTestConditionTimeout.Value) * 1000;

//            bool bStatus = true;

//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();

//            try
//            {
//                //for (int reportIndex = 0; reportIndex < sReport.Length; reportIndex++)
//                //{                    
//                //    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, reportIndex + 2] = sReport[reportIndex];
//                //}

//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5] = sReport;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 6] = s_CableSwIntegrationFinalReportInfo;
//                xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7] = sReport[5];


//                /* Change Color to PASS and Fail */
//                if (sReport == "PASS")
//                {
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
//                }
//                else if (sReport == "FILE")
//                {
//                    /* Write Console Log File as HyperLink in Excel report */
//                    xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7];
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sReport[5], Type.Missing, "ConsoleLog", "ConsoleLog");

//                    if (sComment == "") return;
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);

//                    int iBlank = -1;

//                    for (int i = 8; i < 20; i++)
//                    {
//                        string s = ((Excel.Range)xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, i]).Text;

//                        if (s == null || s == "")
//                        {
//                            iBlank = i;
//                            break;
//                        }                                            
//                    }

//                    if (iBlank > 0)
//                    {
//                        string FileName = Path.GetFileNameWithoutExtension(sComment);
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, iBlank] = sComment;
//                        xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, iBlank];
//                        //xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sComment, Type.Missing, "File", "File");
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sComment, Type.Missing, FileName, "FILE");
//                    }
//                }
//                else if (sReport == "PASSFILE")
//                {
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5] = "PASS";
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

//                    int iBlank = -1;

//                    for (int i = 8; i < 20; i++)
//                    {
//                        string s = ((Excel.Range)xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, i]).Text;
                        
//                        if (s == null || s == "")                            
//                        {
//                            iBlank = i;
//                            break;
//                        }                        
//                    }

//                    if (iBlank > 0)
//                    {
//                        string FileName = Path.GetFileNameWithoutExtension(sComment);
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, iBlank] = sComment;
//                        xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, iBlank];
//                        //xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sComment, Type.Missing, "File", "File");
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sComment, Type.Missing, FileName, "FILE");
//                    }               
                    
//                    //xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 8] = sComment;
//                    //xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 8];
//                    //xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sComment, Type.Missing, "File", "File");
//                }
//                else if (sReport == "FAILFILE")
//                {
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5] = "FAIL";
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

//                    string strTemp = sComment;
//                    string[] s = strTemp.Split(new string[] { "FILE::" }, StringSplitOptions.RemoveEmptyEntries);               
                    
//                    if (s.Length >= 2)
//                    {
//                        xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 6] = s[0];

//                        int iBlank = -1;

//                        for (int i = 8; i < 20; i++)
//                        {
//                            string ss = ((Excel.Range)xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, i]).Text;

//                            if (ss == null || ss == "")
//                            {
//                                iBlank = i;
//                                break;
//                            }                        
//                        }

//                        if (iBlank > 0)
//                        {
//                            s[1] = s[1].Substring(0, s[1].Length -2);
//                            string FileName = Path.GetFileNameWithoutExtension(s[1]);
//                            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, iBlank] = s[1];
//                            xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, iBlank];
//                            //xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sComment, Type.Missing, "File", "File");
//                            xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, s[1], Type.Missing, FileName, "FILE");
//                        }      
                        
//                        //xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 8] = sComment;
//                        //xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 8];
//                        //xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sComment, Type.Missing, "File", "File");
//                    }                    
//                }
//                else
//                {
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
//                }

//                /* Write Console Log File as HyperLink in Excel report */
//                xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7];
//                xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sReport[5], Type.Missing, "ConsoleLog", "ConsoleLog");
//            }
//            catch (Exception ex )
//            {
//                //Debug.WriteLine(" Write Data to Excel File: " + ex.ToString());
//                str = "Write Data to Excel File Exception: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            }

//            try
//            {
//                xls_excelWorkBookRouterChamberPerformanceTest.Save();
//            }
//            catch (Exception ex)
//            {
//                str = "Save Excel Error: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            }
//        }
        
//        /* For reference , This function want't be used */
//        private void RouterChamberPerformanceTestReportData(string[] sReport)
//        {
//            string str = string.Empty;
//            string valueResponsed = string.Empty;
//            string expectedValue = string.Empty;
//            string mibType = string.Empty;
//            int iIndexConsoleLog = 0;
//            string sConsoleLogFile = string.Empty;
//            //s_ErrorMsgForReport = string.Empty;
//            int iConditionTimeout = Convert.ToInt32(nudRouterChamberPerformanceFunctionTestConditionTimeout.Value) * 1000;

//            bool bStatus = true;

//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();

//            try
//            {
//                for (int reportIndex = 0; reportIndex < sReport.Length; reportIndex++)
//                {                    
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, reportIndex + 2] = sReport[reportIndex];
//                }

//                /* Change Color to PASS and Fail */
//                if (sReport == "PASS")
//                {
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
//                }
//                else
//                {
//                    xls_excelWorkSheetRouterChamberPerformanceTest.Range[xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5], xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
//                }

//                /* Write Console Log File as HyperLink in Excel report */
//                xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 7];
//                xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, sReport[5], Type.Missing, "ConsoleLog", "ConsoleLog");
//            }
//            catch (Exception ex )
//            {
//                //Debug.WriteLine(" Write Data to Excel File: " + ex.ToString());
//                str = "Write Data to Excel File Exception: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            }

//            try
//            {
//                xls_excelWorkBookRouterChamberPerformanceTest.Save();
//            }
//            catch (Exception ex)
//            {
//                str = "Save Excel Error: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            }

//        }

//        private string createRouterChamberPerformanceTestSubFolder(string ModelName)
//        {
//            string subFolder = ((ModelName == "") ? "CGA2121_" : ModelName + "_") + DateTime.Now.ToString("yyyyMMdd-HHmmss");

//            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
//                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

//            subFolder = System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder;
//            return subFolder;
//        }

//        private string createRouterChamberPerformanceTestSavePath(string subFolder, ModelInfo info, int Loop)
//        {
//            string PathFile = @"\report\SUBFOLDER\MODEL_SN_SW_HW_MibReadWrite_Test_DATE.csv";

//            PathFile = PathFile.Replace("SUBFOLDER", subFolder);
//            PathFile = PathFile.Replace("MODEL", (info.ModelName == "") ? "CGA2121" : info.ModelName);
//            PathFile = PathFile.Replace("SN", info.SN);
//            PathFile = PathFile.Replace("SW", info.SwVersion);
//            PathFile = PathFile.Replace("HW", info.HwVersion);
//            PathFile = PathFile.Replace("DATE", DateTime.Now.ToString("yyyyMMdd HH-mm"));

//            PathFile = System.Windows.Forms.Application.StartupPath + PathFile;

//            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
//                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

//            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
//                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

//            return PathFile;
//        }

//        private string createRouterChamberPerformanceTestSaveExcelFile(string subFolder, ModelInfo info, int Loop, string FileName)
//        {
//            string PathFile = @"\report\SUBFOLDER\FileName";

//            PathFile = PathFile.Replace("SUBFOLDER", subFolder);
//            PathFile = PathFile.Replace("FileName", FileName);
//            PathFile = PathFile.Replace("DATE", DateTime.Now.ToString("yyyyMMdd HH-mm"));

//            PathFile = System.Windows.Forms.Application.StartupPath + PathFile;

//            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
//                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

//            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
//                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

//            return PathFile;
//        }

//        private string createRouterChamberPerformanceTestSaveExcelFile(string subFolder, string Sku, int Loop, string FileName)
//        {
//            /* string PathFile = @"\report\SUBFOLDER\MODEL_SN_SW_HW_QAMTYPE_LEVELdBmV_(POS)_DATE_(LOOP).xlsx";   */
//            string fileNameWoExt = Path.GetFileNameWithoutExtension(FileName);

//            //string PathFile = @"\report\SUBFOLDER\FileName_SKU_Report_DATE_(LOOP).xlsx";
//            string PathFile = @"SUBFOLDER\FileName_SKU_Report_DATE_(LOOP).xlsx";

//            PathFile = PathFile.Replace("SUBFOLDER", subFolder);
//            PathFile = PathFile.Replace("SKU", Sku);
//            PathFile = PathFile.Replace("FileName", fileNameWoExt);
//            PathFile = PathFile.Replace("DATE", DateTime.Now.ToString("yyyyMMdd HH-mm"));
//            PathFile = PathFile.Replace("LOOP", Loop.ToString());

//            //PathFile = System.Windows.Forms.Application.StartupPath + PathFile;

//            //if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
//            //    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

//            //if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
//            //    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

//            return PathFile;
//        }

//        private string createRouterChamberPerformanceTestVeriwaveSaveExcelFile(string subFolder, int Loop, string FileName)
//        {
//            /* string PathFile = @"\report\SUBFOLDER\MODEL_SN_SW_HW_QAMTYPE_LEVELdBmV_(POS)_DATE_(LOOP).xlsx";   */
//            string fileNameWoExt = Path.GetFileNameWithoutExtension(FileName);

//            string PathFile = @"\report\SUBFOLDER\FileName_Veriwave_Report_DATE_(LOOP).xlsx";

//            PathFile = PathFile.Replace("SUBFOLDER", subFolder);
//            PathFile = PathFile.Replace("FileName", fileNameWoExt);
//            PathFile = PathFile.Replace("DATE", DateTime.Now.ToString("yyyyMMdd HH-mm"));
//            PathFile = PathFile.Replace("LOOP", Loop.ToString());

//            PathFile = System.Windows.Forms.Application.StartupPath + PathFile;

//            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
//                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

//            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
//                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

//            return PathFile;
//        }

//        //private bool CheckNeededParameter_RouterChamberPerformanceTest()
//        //{
//        //    /* Config and Check Test Condition */
//        //    SetText("Check Test Condition...", txtRouterChamberPerformanceFunctionTestInformation);
            
//        //    /* Check Test Condition Content */
//        //    if (dgvCableSwIntegrationExcelTestConditionData.RowCount <= 1)
//        //    {
//        //        MessageBox.Show("Test Condition can't be Empty!!");
//        //        return false;
//        //    }

//        //    /* Check SNMP Parameter */
//        //    if (txtRouterChamberPerformanceFunctionTestCfgFileName.Text == "")
//        //    {
//        //        MessageBox.Show("Cfg File Name can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (txtRouterChamberPerformanceFunctionTestBinFileName.Text == "")
//        //    {
//        //        MessageBox.Show("Bin File Name can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (txtRouterChamberPerformanceFunctionTestCfgHeaderFile.Text == "")
//        //    {
//        //        MessageBox.Show("Cfg Heafer File can't be Empty!!");
//        //        return false;
//        //    }

//        //    if(!File.Exists(txtRouterChamberPerformanceFunctionTestCfgHeaderFile.Text))
//        //    {
//        //        MessageBox.Show("Cfg Header File doesn't Exist!!");
//        //        return false;
//        //    }

//        //    if (txtRouterChamberPerformanceFunctionTestBinHeaderFile.Text == "")
//        //    {
//        //        MessageBox.Show("Bin Heafer File can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (!File.Exists(txtRouterChamberPerformanceFunctionTestBinHeaderFile.Text))
//        //    {
//        //        MessageBox.Show("Bin Header File doesn't Exist!!");
//        //        return false;
//        //    }
            
//        //    if (txtRouterChamberPerformanceFunctionTestCfgConfigConverterFileName.Text == "")
//        //    {
//        //        MessageBox.Show("Cfg Config converter File can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (!File.Exists(txtRouterChamberPerformanceFunctionTestCfgConfigConverterFileName.Text))
//        //    {
//        //        MessageBox.Show("Cfg Config converter File doesn't Exist!!");
//        //        return false;
//        //    }

//        //    if (txtRouterChamberPerformanceFunctionTestBinConfigConverterFileName.Text == "")
//        //    {
//        //        MessageBox.Show("Bin Config converter File can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (!File.Exists(txtRouterChamberPerformanceFunctionTestBinConfigConverterFileName.Text))
//        //    {
//        //        MessageBox.Show("Bin Config converter File doesn't Exist!!");
//        //        return false;
//        //    }

//        //    if (txtRouterChamberPerformanceFunctionTestMibNfsLocation.Text == "")
//        //    {
//        //        MessageBox.Show("NFS Location can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (mtbRouterChamberPerformanceFunctionTestCmIpAddress.Text == "")
//        //    {
//        //        MessageBox.Show("CM IP Address can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (mtbRouterChamberPerformanceFunctionTestMtaIpAddress.Text == "")
//        //    {
//        //        MessageBox.Show("MTA IP Address can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (!CheckIPValid(mtbRouterChamberPerformanceFunctionTestCmIpAddress.Text))
//        //    {
//        //        MessageBox.Show("CM IP Address is Invalid!!");
//        //        return false;
//        //    }

//        //    if (!CheckIPValid(mtbRouterChamberPerformanceFunctionTestMtaIpAddress.Text))
//        //    {
//        //        MessageBox.Show("MTA IP Address is Invalid!!");
//        //        return false;
//        //    }

//        //    //QuickPingRouterChamberPerformanceTest(mtbRouterChamberPerformanceFunctionTestCmIpAddress.Text);

//        //    //QuickPingRouterChamberPerformanceTest(mtbRouterChamberPerformanceFunctionTestMtaIpAddress.Text);

//        //    testConfigRouterChamberPerformanceTest = null;

//        //    /* config file pre-process - convert header file for backup ,especiallly device crash */
//        //    /* CM */
//        //    string str = string.Empty;

//        //    /* Convert default txt to basic cfg file */
//        //    if (File.Exists(s_CableConfigCfgBasicCfgFile))
//        //    {
//        //        SetText("Delete CFG Basic Config File.", txtRouterChamberPerformanceFunctionTestInformation);
//        //        File.Delete(s_CableConfigCfgBasicCfgFile);
//        //    }

//        //    CableConfigFileTxt2Cfg(Path.GetDirectoryName(txtRouterChamberPerformanceFunctionTestCfgConfigConverterFileName.Text), txtRouterChamberPerformanceFunctionTestCfgHeaderFile.Text, s_CableConfigCfgBasicCfgFile, ref str);
//        //    SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

//        //    if (!File.Exists(s_CableConfigCfgBasicCfgFile))
//        //    {
//        //        MessageBox.Show("Convert Basic Config File Failed.");
//        //        return false;
//        //    }

//        //    /* Copy uesr header txt file to txt template */
//        //    if (File.Exists(cableconfigCfgFileTemplate))
//        //    {
//        //        SetText("Delete Header Template File.", txtRouterChamberPerformanceFunctionTestInformation);
//        //        File.Delete(cableconfigCfgFileTemplate);
//        //    }

//        //    File.Copy(txtRouterChamberPerformanceFunctionTestCfgHeaderFile.Text, cableconfigCfgFileTemplate, true);     

//        //    /* Copy txt temp file to txt file */
//        //    File.Copy(cableconfigCfgFileTemplate, cableConfigCfgTxtFile, true);   

//        //    /* MTA */
//        //    //TODO , 
            





//        //     SetText("Check result: OK.", txtRouterChamberPerformanceFunctionTestInformation);
//        //    return true;

//        //    //if (rbtnRouterChamberPerformanceFunctionTestMibExcelFile.Checked)
//        //    //{
//        //    //    // Check Excel File Exist, and Read Excel File
//        //    //    if (txtRouterChamberPerformanceFunctionTestMibExcelFile.Text == "")
//        //    //    {
//        //    //        MessageBox.Show("Excel file can't be Empty!!");
//        //    //        return false;
//        //    //    }

//        //    //    if (txtRouterChamberPerformanceFunctionTestCfgHeaderFile.Text == "")
//        //    //    {
//        //    //        MessageBox.Show("CFG Header file can't be Empty!!");
//        //    //        return false;
//        //    //    }

//        //    //    if (!File.Exists(txtRouterChamberPerformanceFunctionTestMibExcelFile.Text))
//        //    //    {
//        //    //        MessageBox.Show("Excel file doesn't exist!!");
//        //    //        return false;

//        //    //    }

//        //    //    if (!File.Exists(txtRouterChamberPerformanceFunctionTestCfgHeaderFile.Text))
//        //    //    {
//        //    //        MessageBox.Show("CFG Header file doesn't exist!!");
//        //    //        return false;
//        //    //    }

//        //    //    testConfigRouterChamberPerformanceTest = null;

//        //    //    //if (!File.Exists(txtRouterChamberPerformanceFunctionTestMibExcelFile.Text))
//        //    //    //{
//        //    //    //    MessageBox.Show("Excel file doesn't exist!!");
//        //    //    //    return false;
//        //    //    //}

//        //    //    //if (!ReadTestConfig_ReadExcelRouterChamberPerformanceTest(txtRouterChamberPerformanceFunctionTestMibExcelFile.Text))
//        //    //    //{
//        //    //    //    MessageBox.Show("Read Excel file Failed!!");
//        //    //    //    return false;
//        //    //    //}
//        //    //}           
//        //}

//        private bool QuickPingRouterChamberPerformanceTest(string ip, int iTimeout = 3000)
//        {
//            /* Ping if the deivce available*/
//            bool pingStatus = false;
//            //SetText("Ping Device...", txtRouterChamberPerformanceFunctionTestInformation);
//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Reset();
//            stopWatch.Stop();
//            stopWatch.Start();

//            while (true)
//            {
//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    bRouterChamberPerformanceTestFTThreadRunning = false;
//                    //MessageBox.Show("Abort test", "Error");
//                    //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                    //threadRouterChamberPerformanceTestFT.Abort();
//                    return false;
//                    // Never go here ...
//                }


//                if (stopWatch.ElapsedMilliseconds > iTimeout)
//                {
//                    break;
//                }

//                if (PingClient(ip, 1000))
//                {
//                    pingStatus = true;
//                    break;
//                }

//                Thread.Sleep(1000);
//            }

//            if (!pingStatus)
//            {
//                //SetText("Ping Failed!!", txtRouterChamberPerformanceFunctionTestInformation);
//                //MessageBox.Show("Ping Failed, Abort Test!!", "Error");
//                return false;
//            }

//            //SetText("Ping Succeed!!", txtRouterChamberPerformanceFunctionTestInformation);
//            return true;
//        }
        
//        //private void InitialParameter_RouterChamberPerformanceTest()
//        //{
//        //    testConfigRouterChamberPerformanceTest = null;
//        //    xlApp = null;
//        //    xlWorkbook = null;
//        //    xlWorksheet = null;
//        //    xlRange = null;
//        //    st_CbtMibDataConfigReadWriteTest = null;

//        //    txtRouterChamberPerformanceFunctionTestInformation.Text = "";
//        //    return;
//        //}

//        private void ToggleRouterChamberPerformanceFunctionTestGUI()
//        {
//            ToggleRouterChamberPerformanceFunctionTestController(true);
//            Debug.WriteLine("Toggle");
//        }

//        private bool ReadTestConfig_ReadExcelRouterChamberPerformanceTest(string ExcelFile)
//        {
//            int iRealCount = 0; 
//            xlApp = new Excel.Application();
//            xlWorkbook = xlApp.Workbooks.Open(ExcelFile);
//            xlWorksheet = xlWorkbook.Sheets[1];
//            xlRange = xlWorksheet.UsedRange;

//            int rowCount = xlRange.Rows.Count;
//            int colCount = xlRange.Columns.Count;
//            if (rowCount >= 10000) rowCount = 10000;

//            string[,] rowdata = new string[rowCount, colCount]; //Added  for compare index 

//            //iterate over the rows and columns and print to the console as it appears in the file
//            //excel is not zero based!!
//            for (int i = 1; i <= rowCount; i++)
//            {
//                bool bHasData = false;
//                for (int j = 1; j <= colCount; j++)
//                {
//                    //new line
//                    if (j == 1)
//                        Console.Write("\r\n");

//                    //string s1 = xlRange.Cells[i, j]
//                    //string s2 = xlRange.Cells[i, j].Value2;

//                    //write the value to the console
//                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
//                    {
//                        rowdata[i - 1, j - 1] = xlRange.Cells[i, j].Value2.ToString();                        
//                        bHasData = true;
//                    }
//                    else
//                    {
//                        //rowdata[i - 1, j - 1] = "";
//                    }
//                    //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
//                }

//                if (bHasData)
//                    iRealCount++;

//                if (rowdata[i - 1, 0] != null && rowdata[i - 1, 0].ToLower() == "end")
//                    break;
//            }

//            testConfigRouterChamberPerformanceTest = rowdata;

//            //if (rowCount != iRealCount)
//            //{
//            //    string[,] newRowData = new string[iRealCount + 1, colCount];
//            //    for (int i = 0; i < iRealCount; i++)
//            //    {
//            //        for (int j = 0; j < colCount; j++)
//            //        {
//            //            newRowData[i, j] = rowdata[i, j];
//            //        }
//            //    }

//            //    testConfigRouterChamberPerformanceTest = newRowData;
//            //}
//            //else
//            //{
//            //    testConfigRouterChamberPerformanceTest = rowdata;
//            //}

//            /* Close Excel */
//            /* Turn on interactive mode */
//            xlApp.Interactive = true;
//            xlWorkbook.Close();
//            xlApp.Quit();

//            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
//            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
//            xlWorksheet = null;
//            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
//            xlWorkbook = null;
//            releaseExcelObject(xlRange);
//            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
//            xlRange = null;
//            releaseExcelObject(xlApp);
//            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
//            xlApp = null;
//            GC.Collect();

//            return true;
//        }

//        private bool ReadSanityTestConfig_ReadExcelRouterChamberPerformanceTest(string ExcelFile)
//        {
//            int iRealCount = 0;
//            xlApp = new Excel.Application();
//            xlWorkbook = xlApp.Workbooks.Open(ExcelFile);
//            xlWorksheet = xlWorkbook.Sheets[1];
//            xlRange = xlWorksheet.UsedRange;

//            int rowCount = xlRange.Rows.Count;
//            int colCount = xlRange.Columns.Count;
//            if (rowCount >= 10000) rowCount = 10000;

//            string[,] rowdata = new string[rowCount, colCount]; //Added  for compare index 

//            //iterate over the rows and columns and print to the console as it appears in the file
//            //excel is not zero based!!
//            for (int i = 1; i <= rowCount; i++)
//            {
//                bool bHasData = false;
//                for (int j = 1; j <= colCount; j++)
//                {
//                    //new line
//                    if (j == 1)
//                        Console.Write("\r\n");

//                    //string s1 = xlRange.Cells[i, j]
//                    //string s2 = xlRange.Cells[i, j].Value2;

//                    //write the value to the console
//                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
//                    {
//                        rowdata[i - 1, j - 1] = xlRange.Cells[i, j].Value2.ToString();
//                        bHasData = true;
//                    }
//                    else
//                    {
//                        //rowdata[i - 1, j - 1] = "";
//                    }
//                    //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
//                }

//                if (bHasData)
//                    iRealCount++;

//                if (rowdata[i - 1, 0] != null && rowdata[i - 1, 0].ToLower() == "end")
//                    break;
//            }

//            testConfigCableSwIntegrationSanityTest = rowdata;

//            //if (rowCount != iRealCount)
//            //{
//            //    string[,] newRowData = new string[iRealCount + 1, colCount];
//            //    for (int i = 0; i < iRealCount; i++)
//            //    {
//            //        for (int j = 0; j < colCount; j++)
//            //        {
//            //            newRowData[i, j] = rowdata[i, j];
//            //        }
//            //    }

//            //    testConfigCableSwIntegrationSanityTest = newRowData;
//            //}
//            //else
//            //{
//            //    testConfigCableSwIntegrationSanityTest = rowdata;
//            //}

//            /* Close Excel */
//            /* Turn on interactive mode */
//            xlApp.Interactive = true;
//            xlWorkbook.Close();
//            xlApp.Quit();

//            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
//            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
//            xlWorksheet = null;
//            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
//            xlWorkbook = null;
//            releaseExcelObject(xlRange);
//            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
//            xlRange = null;
//            releaseExcelObject(xlApp);
//            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
//            xlApp = null;
//            GC.Collect();

//            return true;
//        }

//        private bool CableChanDirectoryRouterChamberPerformanceTest(string sChannel, Comport cComport)
//        {
//            string str = string.Empty;
//            if (cComport == null)
//            {
//                MessageBox.Show("Comport doesn't Exist!!");
//                return false;
//            }
//            cComport.DiscardBuffer();
//            Thread.Sleep(1000);

//            cComport.WriteLine("");
//            str = cComport.ReadLine();
//            if(str != "")
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            while (true)
//            {
//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    return false;
//                if (str.ToLower().IndexOf("cm/docsisctl>") != -1)
//                    break;
//                else if (str.ToLower().IndexOf("rg>") != -1)
//                    cComport.WriteLine("cd /Console");
//                else if (str.ToLower().IndexOf("rg_console>") != -1)
//                    cComport.WriteLine("sw");
//                //else if (str.ToLower().IndexOf("cm>") != -1)
//                //    cComport.WriteLine("cd /docsis_ctl");
//                else
//                    cComport.WriteLine("cd /docsis_ctl");

//                cComport.WriteLine("");
//                str = cComport.ReadLine();
//                if (str != "")
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                Thread.Sleep(1000);
//            }

//            return true;
//        }




//        private bool RunCableShowCfgRouterChamberPerformanceTest(ref string sInfo, Comport cComport)
//        {
//            string str = string.Empty;
//            if (cComport == null)
//            {
//                MessageBox.Show("Comport doesn't Exist!!");
//                return false;
//            }
//            //cComport.DiscardBuffer();
//            //Thread.Sleep(1000);           

//            bool bCfgStatus = false;
//            str = "";
//            cComport.WriteLine("show cfg");
//            str = cComport.ReadLine();
//            //if (str != "")
//            //{
//            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            //    sInfo += str;
//            //}

//            while (true)
//            {
//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    return false;

//                if (bCfgStatus)
//                {
//                    if (str.ToLower().IndexOf("cm/cm_console/cm>") != -1)
//                    //if (sInfo.ToLower().IndexOf("cm/cm_console/cm>") != -1)
//                        break;
//                }
//                if (str.ToLower().IndexOf("cm config file name") != -1)
//                {
//                    //sInfo += str;
//                    bCfgStatus = true;
//                }

//                if (str != "" && bCfgStatus)
//                {
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    sInfo += str;
//                }

//                str = cComport.ReadLine();                

//                //Thread.Sleep(100);
//            }

//            return true;
//        }




//        private bool CGA2121RebootDut_RouterChamberPerformanceTest()
//        {
//            string str = string.Empty;
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds
//            Stopwatch stopWatch = new Stopwatch();

//            //CGA2121BackToRootDirectory_RouterChamberPerformanceTest("cm");

//            str = "reset Device...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            comPortRouterChamberPerformanceTest.WriteLine("/reset");
//            stopWatch.Start();

//            while (true)
//            {
//                if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value * 1000))
//                {
//                    //str = "Reboot Timeout.";
//                    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    //return false;
//                    break;
//                }

//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    return false;
//                }

//                str = comPortRouterChamberPerformanceTest.ReadLine();

//                if (str != "")
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //if (str.IndexOf(sCableDownloadCfg) != -1)
//                //{
//                //    str = "Reboot Finished.";
//                //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                //    return true;
//                //}
//            }

//            return true;
//        }





//        private bool CGA2121ShowCmCfg_RouterChamberPerformanceTest()
//        {
//            string str = string.Empty;
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds
//            Stopwatch stopWatch = new Stopwatch();            

//            // Check show cfg
//            str = "Check show cfg...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            if (!CGA2121BackToRootDirectory_RouterChamberPerformanceTest("cm"))
//            {
//                return false;
//            }

//            bool bCfgStatus = false;
//            s_CfgContect = "";
//            str = "";
//            comPortRouterChamberPerformanceTest.DiscardBuffer();
//            Thread.Sleep(1000);
//            comPortRouterChamberPerformanceTest.WriteLine("/Console/cm/show cfg");

//            str = comPortRouterChamberPerformanceTest.ReadLine();
//            if (str != "")
//            {
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                s_CfgContect += str;
//            }

//            stopWatch.Stop();
//            stopWatch.Reset();
//            stopWatch.Restart();

//            while (true)
//            {
//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    bRouterChamberPerformanceTestFTThreadRunning = false;
//                    return false;
//                    // Never go here ...
//                }

//                if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//                {
//                    str = "Show cfg Timeout.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    return false;
//                }

//                if (bCfgStatus)
//                {
//                    if (str.ToLower().IndexOf("cm>") != -1)                       
//                        break;
//                }
//                if (str.ToLower().IndexOf("cm config file name") != -1)
//                {
//                    bCfgStatus = true;
//                }

//                if (str != "" && bCfgStatus)
//                {
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    s_CfgContect += str;
//                }

//                str = comPortRouterChamberPerformanceTest.ReadLine();
//                //Thread.Sleep(100);
//            }

//            stopWatch.Stop();
//            stopWatch.Reset();

//            str = "Check Config File Name...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            /* Check if the device loaded cfg file file is equal to CableConfigFile.cfg */
//            string sFwFileName = cableSwIntegrationDutDataSets[i_CableSwIntegrationCurrentDateSetsIndex].CurrentFwName;

//            str = "FW File Name should be: " + sFwFileName;
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            if (s_CfgContect.IndexOf(sFwFileName) == -1)
//            {
//                bConfigFileNameFailed = true;
//                string sErrorMsg = "Config File Name is Wrong!";

//                int iIndexTemp1 = s_CfgContect.ToLower().IndexOf("cm config file name:");
//                int iIndexTemp2 = s_CfgContect.ToLower().IndexOf("cm config file contents:");

//                if (iIndexTemp1 >= 0 && iIndexTemp2 >= 0)
//                {
//                    sErrorMsg = s_CfgContect.Substring(iIndexTemp2, iIndexTemp2 - iIndexTemp1);
//                }

//                str = "Error Config File Name: " + sErrorMsg;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                return false;
//            }

//            return true;            
//        }

//        private bool CGA2121CheckFwFileName_RouterChamberPerformanceTest(int iIndex)
//        {
//            string str = string.Empty;
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds
//            Stopwatch stopWatch = new Stopwatch();

//            // Check show cfg
//            str = "Show Firmware File Name...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //if (!CGA2121BackToRootDirectory_RouterChamberPerformanceTest("cm"))
//            //{
//            //    return false;
//            //}

//            bool bCfgStatus = false;
//            s_CfgContect = "";
//            str = "";
//            comPortRouterChamberPerformanceTest.DiscardBuffer();
//            Thread.Sleep(1000);

//            comPortRouterChamberPerformanceTest.WriteLine("/ver");

//            str = comPortRouterChamberPerformanceTest.ReadLine();
//            if (str != "")
//            {
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                s_CfgContect += str;
//            }

//            stopWatch.Stop();
//            stopWatch.Reset();
//            stopWatch.Restart();

//            while (true)
//            {
//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    bRouterChamberPerformanceTestFTThreadRunning = false;
//                    return false;
//                    // Never go here ...
//                }

//                if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//                {
//                    str = "Show version Timeout.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    return false;
//                }

//                if (bCfgStatus)
//                {
//                    if (str.ToLower().IndexOf(">") != -1)
//                        break;
//                }
//                if (str.ToLower().IndexOf("version:") != -1)
//                {
//                    bCfgStatus = true;
//                    s_CfgContect += str;
//                    comPortRouterChamberPerformanceTest.WriteLine("");
//                    //break;
//                }

//                if (str != "" && bCfgStatus)
//                {
//                    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    s_CfgContect += str;
//                }

//                str = comPortRouterChamberPerformanceTest.ReadLine();
//                if (str != "")
//                {
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                }
//            }

//            stopWatch.Stop();
//            stopWatch.Reset();

//            str = "Check Firmware File Name...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            /* Check if the device loaded cfg file file is equal to CableConfigFile.cfg */
//            string sFwFileName = cableSwIntegrationDutDataSets[iIndex].CurrentFwName;

//            str = "FW File Name should be: " + sFwFileName;
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            sFwFileName = sFwFileName.Substring(0, sFwFileName.Length - 4);

//            if (s_CfgContect.IndexOf(sFwFileName) <0 )
//            {
//                bConfigFileNameFailed = true;
//                string sErrorMsg = "Firmware File Name is Wrong!";

//                int iIndexTemp1 = s_CfgContect.ToLower().IndexOf("version:");
//                int iIndexTemp2 = s_CfgContect.ToLower().IndexOf("codec");

//                if (iIndexTemp1 >= 0 && iIndexTemp2 >= 0)
//                {
//                    sErrorMsg = s_CfgContect.Substring(iIndexTemp2, iIndexTemp2 - iIndexTemp1);
//                }

//                str = "Error Firmware File Name: " + sErrorMsg;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                return false;
//            }

//            return true;
//        }

//        private bool CGA2121RebootBySnmpReadByComport_RouterChamberPerformanceTest()
//        {
//            string str = string.Empty;
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds
//            Stopwatch stopWatch = new Stopwatch();

//            CGA2121BackToRootDirectory_RouterChamberPerformanceTest("cm");

//            str = "reset Device...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            comPortRouterChamberPerformanceTest.WriteLine("/reset");
//            stopWatch.Start();

//            while (true)
//            {
//                if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value * 1000))
//                {
//                    str = "Reboot Timeout.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    return false;
//                }

//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    return false;
//                }

//                str = comPortRouterChamberPerformanceTest.ReadLine();

//                if (str != "")
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //if (str.IndexOf(sCableDownloadCfg) != -1)
//                //{
//                //    str = "Reboot Finished.";
//                //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                //    return true;
//                //}
//            }

//            return true;
//        }
        
//        private bool RestoreBasicConfigFileAndRebootCableByConsoleRouterChamberPerformanceTest()
//        {
//            string str = string.Empty;
//            Stopwatch stopWatch = new Stopwatch();

//            //Copy Basic Config File to tftp Server
//            //if(s_TestType.ToLower() == "cm")
//            //{
//            //    RestoreCfgBasicConfigFileRouterChamberPerformanceTest();
//            //    //Reboot DUT

//            //}

            
//            ChangeToCableConsoleCM(ref str, comPortRouterChamberPerformanceTest);
//            comPortRouterChamberPerformanceTest.WriteLine("/reset");
//            stopWatch.Start();            

//            while (true)
//            {
//                if(stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value *1000))
//                {
//                    str = "Reboot Timeout.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    return false;
//                }

//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    return false;
//                }

//                str = comPortRouterChamberPerformanceTest.ReadLine();

//                if(str != "") 
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                if (str.IndexOf(sCableDownloadCfg) != -1)
//                {
//                    str = "Reboot Finished.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    return true;                        
//                }
//            }
//        }

//        private bool RestoreCfgBasicConfigFileRouterChamberPerformanceTest()
//        {
//            string str = string.Empty;
//            string NfsPath = "";// txtRouterChamberPerformanceFunctionTestMibNfsLocation.Text;
//            string sBasicCfgFile = "";//txtRouterChamberPerformanceFunctionTestCfgFileName.Text;

//            //File.Copy(s_CableConfigCfgBasicCfgFile, cableConfigBinFile, true);
//            File.Copy(s_CableConfigCfgBasicCfgFile, sBasicCfgFile, true);
            
//            str = " Copy basic cfg file to tftp server.";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            CopyFileToNfsFoler(sBasicCfgFile, NfsPath, true);
//            str = "Copy File Succeed.";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            return true;
//        }

        private bool ParameterCheckAndInitial_RouterChamberPerformanceTest()
        {
            /* Initialize  */
//            testConfigRouterChamberPerformanceTest = null;
//            xlApp = null;
//            xlWorkbook = null;
//            xlWorksheet = null;
//            xlRange = null;
//            st_CbtMibDataConfigReadWriteTest = null;

//            txtRouterChamberPerformanceFunctionTestInformation.Text = "";
//            bCableSwIntegrationThroughputTest = false;
//            t_Timer = null;
//            Snmp_RouterChamberPerformanceTest = null;

//            //i_SetCheckHourTime = 5;
//            //i_SetCheckMinuteTime = 00;
//            //i_SetCheckSecondTime = 00;

//            /* Check the day now is less than the end day */
//            if (DateTime.Compare(dtpCableSwIntegrationConfigurationFtpTestPeriodStartDay.Value, dtpCableSwIntegrationConfigurationFtpTestPeriodEndDay.Value) > 0)
//            {
//                MessageBox.Show("Start Day can't be late than End Day!!");
//                return false;
//            }

//            if (DateTime.Compare(DateTime.Now, dtpCableSwIntegrationConfigurationFtpTestPeriodEndDay.Value) > 0)
//            {
//                MessageBox.Show("Today is not in the Test Period!!");
//                return false;
//            }

//            /* Config and Check Test Condition */
//            SetText("Check Paramters...", txtRouterChamberPerformanceFunctionTestInformation);

//            /* Check Test Condition Content */
//            if (dgvRouterChamberPerformanceDutsSettingData.RowCount <= 1)
//            {
//                MessageBox.Show("Device Setting Content can't be Empty!!");
//                return false;
//            }
            
//            /* Check SNMP Parameter */
//            if (mtbCableSwIntegrationConfigurationCmtsCasaC10gIpAddress.Text == "" || mtbCableSwIntegrationConfigurationCmtsArrisC4IpAddress.Text == "")
//            {
//                MessageBox.Show("CMTS server IP Address can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationRdFtpServerPath.Text == "" || txtCableSwIntegrationConfigurationNfsServerPath.Text == "")
//            {
//                MessageBox.Show("Ftp server Path can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationReportEmailSendTo.Text == "")
//            {
//                MessageBox.Show("Email Send to list can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationTestConditionExcelFile.Text == "")
//            {
//                MessageBox.Show("Test Condition Excel File Name can't be Empty!!");
//                return false;
//            }            

//            if (!CheckIPValid(mtbCableSwIntegrationConfigurationCmtsCasaC10gIpAddress.Text) || !CheckIPValid(mtbCableSwIntegrationConfigurationCmtsArrisC4IpAddress.Text))
//            {
//                MessageBox.Show("CMTS Server IP Address is Invalid!!");
//                return false;
//            }

//            //Device SNMP Content Check 
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Description OID can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationResetOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Reset OID can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationResetValue.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Reset Value can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationTftpServerOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Tftp Server OID can't be Empty!!");
//                return false;
//            }

//            if (mtbCableSwIntegrationConfigurationServerIp.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Tftp Server IP Value can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationImageFileOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: FW Image File OID can't be Empty!!");
//                return false;
//            }

//            if (dudCableSwIntegrationConfigurationResetType.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Reset Type can't be Empty!!");
//                return false;
//            }
            
//            if (dudCableSwIntegrationConfigurationAdminStatusType.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Admin Status Type can't be Empty!!");
//                return false;
//            }

//            if (dudCableSwIntegrationConfigurationTftpServerTypeType.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Tftp Server Type Type can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationAdminStatusOID.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Admin Status OID can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationAdminStatusValue.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Admin Status Value can't be Empty!!");
//                return false;
//            }
            
//            if (txtCableSwIntegrationConfigurationTftpServerTypeOID.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Tftp Server Type OID can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationTftpServerTypeValue.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Tftp Server Type Value can't be Empty!!");
//                return false;
//            }

//            if (nudCableSwIntegrationConfigurationInProcessTimeout.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: In-Process Timeout Value can't be Empty!!");
//                return false;
//            }
            
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }

//            if (nudCableSwIntegrationConfigurationFwUpdateWaitTime.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Wait FW Download Time Value can't be Empty!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationCfgConverterFileName.Text == "")
//            {
//                MessageBox.Show("Configuration Cfg Converter File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtCableSwIntegrationConfigurationCfgConverterFileName.Text))
//            {
//                MessageBox.Show("Configuration Cfg Converter File doesn't exist!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationCfgSnmpEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("Snmp Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtCableSwIntegrationConfigurationCfgSnmpEuTxtFileName.Text))
//            {
//                MessageBox.Show("Snmp Euro Txt File doesn't exist!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationCfgTlvEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("Tlv Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtCableSwIntegrationConfigurationCfgTlvEuTxtFileName.Text))
//            {
//                MessageBox.Show("Tlv Euro Txt File doesn't exist!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationCfgP7bEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("P7b Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtCableSwIntegrationConfigurationCfgP7bEuTxtFileName.Text))
//            {
//                MessageBox.Show("P7b Euro Txt doesn't exist!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationCfgSnmpNonEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("Snmp Non-Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtCableSwIntegrationConfigurationCfgSnmpNonEuTxtFileName.Text))
//            {
//                MessageBox.Show("Snmp Non-Euro Txt File doesn't exist!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationCfgTlvNonEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("Tlv Non-Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtCableSwIntegrationConfigurationCfgTlvNonEuTxtFileName.Text))
//            {
//                MessageBox.Show("Tlv Non-Euro Txt File doesn't exist!!");
//                return false;
//            }

//            if (txtCableSwIntegrationConfigurationCfgP7bNonEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("P7b Non-Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtCableSwIntegrationConfigurationCfgP7bNonEuTxtFileName.Text))
//            {
//                MessageBox.Show("P7b Non-Euro Txt File doesn't exist!!");
//                return false;
//            }
            
//            SetText("Done!!", txtRouterChamberPerformanceFunctionTestInformation);
//            SetText("Read DUT Settings...", txtRouterChamberPerformanceFunctionTestInformation);

//            /* Read Device Data Setting */
//            ReadDataSets_RouterChamberPerformanceTest();

//            if (bCableSwIntegrationThroughputTest)
//            {
//                //if (btnCableSwIntegrationVeriwaveTestConditionEditSetting.Text.ToLower() == "cancel")
//                //{
//                //    btnCableSwIntegrationVeriwaveTestConditionEditSetting.Text = "Edit";
//                //    hasDeleteButton = false;
//                //    dgvCableSwIntegrationVeriwaveTestConditionData.Columns.Remove("Action");
//                //}

//                //if (dgvCableSwIntegrationVeriwaveTestConditionData.RowCount > 1)
//                //{
//                //    //if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
//                //    //    return;
//                //    //else
//                //    //{
//                //    //    DataTable dt = (DataTable)dgvCableSwIntegrationVeriwaveTestConditionData.DataSource;
//                //        dgvCableSwIntegrationVeriwaveTestConditionData.Rows.Clear();
//                //        //dgvCableSwIntegrationVeriwaveTestConditionData.DataSource = dt;
//                //    //}
//                //}
//                //readXmlCableSwIntegrationVeriwaveTestCondition(System.Windows.Forms.Application.StartupPath + "\\testCondition\\CableSwIntegrationVeriwaveTestCondition.xml");

//                if (txtCableSwIntegrationConfigurationWaveappFolder.Text == "")
//                {
//                    MessageBox.Show("Veriwave Wave app Folder can't be Empty!!");
//                    return false;
//                }

//                if (!File.Exists(txtCableSwIntegrationConfigurationWaveappFolder.Text))
//                {
//                    MessageBox.Show("Veriwave Wave app File doesn't Exist!!");
//                    return false;
//                }

//                if (txtCableSwIntegrationConfigurationReportFolder.Text == "")
//                {
//                    MessageBox.Show("Veriwave Rrport Folder can't be Empty!!");
//                    return false;
//                }

//                if (!Directory.Exists(txtCableSwIntegrationConfigurationReportFolder.Text))
//                {
//                    try
//                    {
//                        string[] saDirTemp = txtCableSwIntegrationConfigurationReportFolder.Text.Split('\\');
//                        string sDir = saDirTemp[0] + "\\" + saDirTemp[1];
//                        if (!Directory.Exists(sDir)) Directory.CreateDirectory(sDir);

//                        for (int i = 2; i < saDirTemp.Length; i++)
//                        {
//                            sDir += "\\" + saDirTemp[i];
//                            if (!Directory.Exists(sDir)) Directory.CreateDirectory(sDir);
//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        MessageBox.Show("Create Veriwave report Folder Faided!!");
//                        return false;
//                    }
//                }

//                if (!Directory.Exists(txtCableSwIntegrationConfigurationReportFolder.Text))
//                {
//                    MessageBox.Show("Veriwave Report Folder doesn't exist!!");
//                    return false;
//                }

//                if (dgvCableSwIntegrationVeriwaveTestConditionData.RowCount <= 1)
//                {
//                    MessageBox.Show("Veriwave Test Condition can't be Empty!!");
//                    return false;
//                }

//                /* Set Veriwave app cl foleder */
//                string waveFile = txtCableSwIntegrationConfigurationWaveappFolder.Text;
//                string wavePath = waveFile.Substring(0, waveFile.Length - Path.GetFileName(waveFile).Length);
//                s_WaveappclFile = wavePath + "waveapps_cl.exe";

//                /* Check waveapps_cl.exe */
//                //SetText("Check waveapps_cl.exe...", txtRouterChamberPerformanceFunctionTestInformation);
//                if (!File.Exists(s_WaveappclFile))
//                {
//                    MessageBox.Show("waveapps_cl.exe doesn't exist!!");
//                    return false;
//                }

//                sa_VeriwaveOrigReportFile = null;
//                s_VeriwaveReportFolder = txtCableSwIntegrationConfigurationReportFolder.Text;
//                string[] sa_temp;
//                GetDirectoryFolders_CommonFunction(txtCableSwIntegrationConfigurationReportFolder.Text, out sa_temp);
//                sa_VeriwaveOrigReportFile = new string[sa_temp.Length + dgvCableSwIntegrationVeriwaveTestConditionData.RowCount];
//                for (int i = 0; i < sa_temp.Length; i++)
//                {
//                    sa_VeriwaveOrigReportFile[i] = sa_temp[i];
//                }

//                //SetText("Check Succeed!", txtRouterChamberPerformanceFunctionTestInformation);
//            }         
            
//            /* Read Test Condition */
//            string sExcelFileName = txtCableSwIntegrationConfigurationTestConditionExcelFile.Text;
//            string str = string.Empty;

//            if (!File.Exists(txtCableSwIntegrationConfigurationTestConditionExcelFile.Text))
//            {
//                MessageBox.Show("Test Condition Excel File doesn't Exist!!");
//                return false;
//            }

//            if (!ReadSanityTestConfig_ReadExcelRouterChamberPerformanceTest(sExcelFileName))
//            {
//                MessageBox.Show("Read Excel Failed!!");
//                return false;
//            }            

//            sExcelFileName = System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterChamberPerformanceTestSanityPreVersion.xlsx";

//            if (!File.Exists(sExcelFileName))
//            {
//                MessageBox.Show("Test Condition Excel File doesn't Exist!!");
//                return false;
//            }

//            if (!ReadTestConfig_ReadExcelRouterChamberPerformanceTest(sExcelFileName))
//            {
//                MessageBox.Show("Read Excel Failed!!");
//                return false;
//            }

//            s_EmailSenderID = txtCableSwIntegrationConfigurationReportEmailSenderGmailAccount.Text;
//            s_EmailSenderPassword = txtCableSwIntegrationConfigurationReportEmailSenderGmailPassword.Text;
//            s_EmailReceivers = txtCableSwIntegrationConfigurationReportEmailSendTo.Text.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            
//            SetText("Done!!", txtRouterChamberPerformanceFunctionTestInformation);

//            if (!ReadDeviceSsidSecurityModePresharedKey_RouterChamberPerformanceTest())
//            {
//                //MessageBox.Show("Read SSID Security Key or Preshared Key Failed!!");
//                return false;
//            }

            return true;
        }

//        private bool ReadDataSets_RouterChamberPerformanceTest()
//        {
//            CableSwIntegrationDutDataSets[] sDataSets = new CableSwIntegrationDutDataSets[dgvRouterChamberPerformanceDutsSettingData.RowCount - 1];
//            //string[,] dataSets = new string[dgvRouterChamberPerformanceDutsSettingData.RowCount - 1, dgvRouterChamberPerformanceDutsSettingData.ColumnCount];

//            for (int i = 0; i < dgvRouterChamberPerformanceDutsSettingData.RowCount - 1; i++)
//            {
//                sDataSets[i].index = Convert.ToInt32(dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[0].Value.ToString());
//                sDataSets[i].IpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[1].Value.ToString();
//                sDataSets[i].MacAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[2].Value.ToString();
//                sDataSets[i].PcIpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[3].Value.ToString();
//                sDataSets[i].SwitchPort = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[4].Value.ToString();
//                sDataSets[i].ComportNum = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[5].Value.ToString();
//                sDataSets[i].TestType = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[6].Value.ToString();
//                sDataSets[i].CmtsType = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[7].Value.ToString();
//                sDataSets[i].SkuModel = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[8].Value.ToString();
//                sDataSets[i].SerialNumber = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[9].Value.ToString();
//                sDataSets[i].SwVersion = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[10].Value.ToString();
//                sDataSets[i].HwVersion = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[11].Value.ToString();
//                sDataSets[i].FwFileName = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[12].Value.ToString();
//                sDataSets[i].SSID = dgvRouterChamberPerformanceDutsSettingData.Rows[i].Cells[13].Value.ToString();
//            }

//            cableSwIntegrationDutDataSets = sDataSets;

//            int snmpVersion = rbtnCableSwIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnCableSwIntegrationConfigurationSnmpV2.Checked ? 2 : 3;

//            try
//            {
//                for (int i = 0; i < cableSwIntegrationDutDataSets.Length; i++)
//                {
//                    cableSwIntegrationDutDataSets[i].snmp = new CBTSnmp();
//                    cableSwIntegrationDutDataSets[i].snmp.Init(cableSwIntegrationDutDataSets[i].IpAddress, snmpVersion, dudCableSwIntegrationConfigurationSnmpReadCommunity.Text, dudCableSwIntegrationConfigurationSnmpWriteCommunity.Text, Convert.ToInt32(nudCableSwIntegrationConfigurationSnmpPort.Value), Convert.ToInt32(nudCableSwIntegrationConfigurationTrapPort.Value));
//                }
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show("Snmp Initial Failed!!: " + ex.ToString());
//                System.Windows.Forms.Cursor.Current = Cursors.Default;
//                btnRouterChamberPerformanceFunctionTestRun.Enabled = true;
//                return false;
//            }

//            foreach (CableSwIntegrationDutDataSets set in sDataSets)
//            {
//                if (set.TestType.ToLower().IndexOf("throughput") >= 0)
//                {
//                    bCableSwIntegrationThroughputTest = true;                    
//                }
//            }            

//            return true;
//        }

//        private bool ReadDeviceSsidSecurityModePresharedKey_RouterChamberPerformanceTest()
//        {
//            string str = string.Empty;
//            SetText("Read SSID, Security Mode, Wpa PreShared key before Test!!", txtRouterChamberPerformanceFunctionTestInformation);

//            /* Read TchRgDot11BssSsid , TchRgDot11BssSecurityMode , TchRgDot11WpaPreSharedkey Before test */
//            /* OID: TchRgDot11BssSsid , TchRgDot11BssSecurityMode , TchRgDot11WpaPreSharedkey */
//            //const string cs_CGA2121BssSsid24gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.2.1.1.3.32";
//            //const string cs_CGA2121BssSsid5gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.2.1.1.3.112";
//            //const string cs_CGA2121BssSecurityMode24gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.2.1.1.4.32";
//            //const string cs_CGA2121BssSecurityMode5gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.2.1.1.4.112";
//            //const string cs_CGA2121WpaPreSharedkey24gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.3.1.1.2.32";
//            //const string cs_CGA2121WpaPreSharedkey5gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.3.1.1.2.112";

//            for (int i = 0; i < cableSwIntegrationDutDataSets.Length; i++)
//            {
//                //if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                //{
//                //    //MessageBox.Show("Abort!!");
//                //    return false;
//                //}

//                string sIp = cableSwIntegrationDutDataSets[i].IpAddress;
//                str = "Ping Device First: " + sIp;
//                SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

//                ////Ping Device First
//                ///* Ping ip first */
//                ////if (!QuickPingRouterChamberPerformanceTest(mtbRouterChamberPerformanceFunctionTestCmIpAddress.Text, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//                //if (!QuickPingRouterChamberPerformanceTest(sIp, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//                //{
//                //    str = String.Format("Sku: {0} Ping Device Failed.", cableSwIntegrationDutDataSets[i].SkuModel);
//                //    SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

//                //    MessageBox.Show("Ping Device Failed!!");
//                //    return false;
//                //}

//                /* Ping if the deivce available*/
//                bool pingStatus = false;
//                int iTimeout = 5 * 1000; //5 Seconds
//                //SetText("Ping Device...", txtRouterChamberPerformanceFunctionTestInformation);
//                Stopwatch stopWatch = new Stopwatch();
//                stopWatch.Reset();
//                stopWatch.Stop();
//                stopWatch.Start();

//                while (true)
//                {
//                    //if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    //{
//                    //    bRouterChamberPerformanceTestFTThreadRunning = false;
//                    //    //MessageBox.Show("Abort test", "Error");
//                    //    //this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                    //    //threadRouterChamberPerformanceTestFT.Abort();
//                    //    return false;
//                    //    // Never go here ...
//                    //}


//                    if (stopWatch.ElapsedMilliseconds > iTimeout)
//                    {
//                        break;
//                    }

//                    if (PingClient(sIp, 1000))
//                    {
//                        pingStatus = true;
//                        break;
//                    }

//                    Thread.Sleep(1000);
//                }

//                if (!pingStatus)
//                {
//                    SetText("Ping Failed!!", txtRouterChamberPerformanceFunctionTestInformation);
//                    str = String.Format("Sku: {0}, IP:{1} Ping Failed!!", cableSwIntegrationDutDataSets[i].SkuModel, sIp);
//                    MessageBox.Show(str, "Error");
//                    return false;
//                }

//                SetText("Ping Succeed!!", txtRouterChamberPerformanceFunctionTestInformation);

//                //if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                //{
//                //    //MessageBox.Show("Abort!!");
//                //    return false;
//                //}

//                string valueResponsed = string.Empty;
//                string mibType = string.Empty;
//                int snmpVersion = rbtnCableSwIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnCableSwIntegrationConfigurationSnmpV2.Checked ? 2 : 3;

//                try
//                {
//                    /* 2.4G SSID */
//                    str = String.Format("Sku: {0} 2.4G SSID.", cableSwIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121BssSsid24gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} SSID is not Valid.", cableSwIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    cableSwIntegrationDutDataSets[i].PreSsid24G = valueResponsed.Trim();

//                    //if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    //{
//                    //    //MessageBox.Show("Abort!!");
//                    //    return false;
//                    //}

//                    /* 5G SSID */
//                    str = String.Format("Sku: {0} 5G SSID.", cableSwIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121BssSsid5gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} SSID is not Valid.", cableSwIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    cableSwIntegrationDutDataSets[i].PreSsid5G = valueResponsed.Trim();

//                    //if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    //{
//                    //    //MessageBox.Show("Abort!!");
//                    //    return false;
//                    //}

//                    /* 2.4G Security Mode */
//                    str = String.Format("Sku: {0} 2.4G Security Mode.", cableSwIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121BssSecurityMode24gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} Security Mode is not Valid.", cableSwIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    cableSwIntegrationDutDataSets[i].PreSecurityMode24G = valueResponsed.Trim();

//                    //if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    //{
//                    //    //MessageBox.Show("Abort!!");
//                    //    return false;
//                    //}

//                    /* 5G Security Mode */
//                    str = String.Format("Sku: {0} 5G Security Mode.", cableSwIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121BssSecurityMode5gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} Security Mode is not Valid.", cableSwIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    cableSwIntegrationDutDataSets[i].PreSecurityMode5G = valueResponsed.Trim();

//                    //if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    //{
//                    //    //MessageBox.Show("Abort!!");
//                    //    return false;
//                    //}

//                    /* 2.4G PreShared Key */
//                    str = String.Format("Sku: {0} 2.4G PreShared Key.", cableSwIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121WpaPreSharedkey24gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} PreShared Key is not Valid.", cableSwIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    cableSwIntegrationDutDataSets[i].PreWpaPreSharedkey24G = valueResponsed.Trim();

//                    //if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    //{
//                    //    //MessageBox.Show("Abort!!");
//                    //    return false;
//                    //}

//                    /* 5G PreShared Key */
//                    str = String.Format("Sku: {0} 5G PreShared Key.", cableSwIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterChamberPerformanceFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121WpaPreSharedkey5gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} PreShared Key is not Valid.", cableSwIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    cableSwIntegrationDutDataSets[i].PreWpaPreSharedkey5G = valueResponsed.Trim();
//                }
//                catch (Exception ex)
//                {
//                    str = String.Format("Sku: {0} Read SSId, Security Mode or PreShared Key Exception: {1}", cableSwIntegrationDutDataSets[i].SkuModel, ex.ToString());
//                    SetText(str, txtRouterChamberPerformanceFunctionTestInformation);
//                    MessageBox.Show(str);
//                    return false;
//                }
//            }

//            SetText("Done!!", txtRouterChamberPerformanceFunctionTestInformation);

//            return true;
//        }
        
//        private bool CGA2121BackToRootDirectory_RouterChamberPerformanceTest(string mode)
//        {
//            string str = string.Empty;
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds

//            /* Create a time check */
//            Stopwatch stopWatch = new Stopwatch();

//            comPortRouterChamberPerformanceTest.DiscardBuffer();
//            Thread.Sleep(1000);

//            comPortRouterChamberPerformanceTest.WriteLine(" cd /");
//            comPortRouterChamberPerformanceTest.WriteLine("");
//            str = comPortRouterChamberPerformanceTest.ReadLine();
//            if (str != "")
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
            
//            stopWatch.Stop();
//            stopWatch.Reset();
//            stopWatch.Restart();

//            while (true)
//            {
//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                    return false;

//                if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//                {
//                    str = "Change Directory Timeout.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    return false;
//                }

//                if (mode.ToLower() == "cm")
//                {
//                    if (str.ToLower().IndexOf("cm>") >=0 ) break;

//                    if (str.ToLower().IndexOf("rg>") <0)
//                    {
//                        comPortRouterChamberPerformanceTest.WriteLine("/Console/sw");
//                        str = comPortRouterChamberPerformanceTest.ReadLine();
//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        Thread.Sleep(3000);
//                        comPortRouterChamberPerformanceTest.WriteLine("");
//                        str = comPortRouterChamberPerformanceTest.ReadLine();
//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        comPortRouterChamberPerformanceTest.WriteLine("cd /");
//                    }
//                }

//                if (mode.ToLower() == "mta")
//                {
//                    if (str.ToLower().IndexOf("rg>") >=0) break;

//                    if (str.ToLower().IndexOf("cm>") <0)
//                    {
//                        comPortRouterChamberPerformanceTest.WriteLine("/Console/sw");
//                        str = comPortRouterChamberPerformanceTest.ReadLine();
//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        Thread.Sleep(3000);
//                        comPortRouterChamberPerformanceTest.WriteLine("");
//                        str = comPortRouterChamberPerformanceTest.ReadLine();
//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        comPortRouterChamberPerformanceTest.WriteLine("cd /");
//                    }
//                }

//                comPortRouterChamberPerformanceTest.WriteLine("");
//                str = comPortRouterChamberPerformanceTest.ReadLine();
//            }          

//            return true;
//        }

//        private bool CreateSkuFoldRouterChamberPerformanceTest()
//        {
//             bool isExists;
//             string subPath = m_RouterChamberPerformanceTestSubFolder;

//            isExists = System.IO.Directory.Exists(subPath);
//            if (!isExists)
//                System.IO.Directory.CreateDirectory(subPath);

//            //if (bCableSwIntegrationThroughputTest)
//            //{
//            //    subPath = subPath + "\\Veriwave";
//            //    cableSwIntegrationVeriwave.ReportSavePath = subPath;
//            //    isExists = System.IO.Directory.Exists(subPath);
//            //    if (!isExists)
//            //        System.IO.Directory.CreateDirectory(subPath);
//            //}

//            for (int i = 0; i < cableSwIntegrationDutDataSets.Length; i++)
//            {
//                subPath = subPath + "\\" + cableSwIntegrationDutDataSets[i].SkuModel;
//                cableSwIntegrationDutDataSets[i].ReportSavePath = subPath;
//                isExists = System.IO.Directory.Exists(subPath);
//                if (!isExists)
//                    System.IO.Directory.CreateDirectory(subPath);
//            }
            
//            return true;
//        }

//        private bool CreateSubFoldCableSwIntegrationVeriwave()
//        {
//            bool isExists;
//            string subPath = m_RouterChamberPerformanceTestSubFolder;

//            isExists = System.IO.Directory.Exists(subPath);
//            if (!isExists)
//                System.IO.Directory.CreateDirectory(subPath);

//            subPath = subPath + "\\Veriwave";
//            cableSwIntegrationVeriwave.ReportSavePath = subPath;
//            isExists = System.IO.Directory.Exists(subPath);
//            if (!isExists)
//                System.IO.Directory.CreateDirectory(subPath);
           
//            return true;
//        }
        
//        private bool CopyRdFwToNfsServerRouterChamberPerformanceTest()
//        {
//            string str = string.Empty;
//            //string[] sSku = new string[] { "CHILE", "COLOMBIA", "PERU", "GERMANY" };
//            List<string> list = new List<string>();
//            bool bRdFwDownloadStatus = true;
//            string[] sInTempFile = new string[30];

//            string sDay = DateTime.Now.ToString("yyyy-MM-dd");
//            string sShorDay = DateTime.Now.ToString("yyMMdd");

//            //string RdPath = "192.168.65.201";
//            string RdPath = txtCableSwIntegrationConfigurationRdFtpServerPath.Text;
//            string RdTempPath = System.Windows.Forms.Application.StartupPath + "\\Temp";
//            string sDirOrig = "/Technicolor/Taipei_BFC5.7.1mp3_RG/CGA2121";
//            string sDir = sDirOrig;
//            //string sFwFileNameOid = ".1.3.6.1.2.1.69.1.3.2.0";
//            //string sFwFileNameOidType = "OctetString";
//            //string[] sFwFiles = new string[cableSwIntegrationDutDataSets.Length];
            

//            str = "Date is " + sDay;
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            for (int i = 0; i < cableSwIntegrationDutDataSets.Length; i++)
//            {
//                sDir = sDirOrig;
//                //sDir = sDir + "/CGA2121_" + sSku[i] + "/";
//                if (cableSwIntegrationDutDataSets[i].SkuModel.IndexOf("GERMANY") >= 0)
//                {
//                    sDir = sDir + "/CGA2121_GERMANY/";
//                }
//                else
//                {
//                    sDir = sDir + "/CGA2121_" + cableSwIntegrationDutDataSets[i].SkuModel + "/";
//                }
//                sDir = RdPath + sDir;

//                str = "Ftp Path is: ftp://" + sDir;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                CbtFtpClient ftpclient = new CbtFtpClient(sDir, "cbtsqa", "jenkins");

//                try
//                {
//                    /* Check if the folder exists */
//                    list = ftpclient.GetFtpDirList();
//                    bool bCheckPoint = false;

//                    foreach (string s in list)
//                    {                        
//                        if (s == sDay)
//                        {
//                            sDir += sDay + "/";
//                            bCheckPoint = true;                           
//                            break;
//                        }
//                    }

//                    if (!bCheckPoint)
//                    { //找不到今天日期的資料夾,等一個小時之後再試
//                        str = "The Date Folder doesn't exist!! ftp://" + sDir + sDay;
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        bRdFwDownloadStatus = false;                        

//                        //超過早上10點都沒有FW,就隔天再做.
//                        if (i_SetCheckHourTime >= 10)
//                        {
//                            i_SetCheckHourTime = 5;

//                            str = "Try again Tomorrow morning!!!";
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                            i_SetCheckHourTime++;

//                            str = "Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                            return false;
//                        }

//                        str = "Wait for one hour and try again!!!";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        i_SetCheckHourTime++;

//                        str = "Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        return false;
//                    }                   

//                    str = "The Date Folder Exists: ftp://" + sDir;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    ftpclient = null;
//                    ftpclient = new CbtFtpClient(sDir, "cbtsqa", "jenkins");

//                    /* Get File list in today's folder */
//                    list.Clear();
//                    list = ftpclient.GetFtpFileList();
//                    string[] sFileName = new string[20];
//                    //iIndex = 0;
//                    string sDownloadFileName = cableSwIntegrationDutDataSets[i].FwFileName;
//                    sDownloadFileName = sDownloadFileName.Replace("#DATE", sShorDay);
//                    cableSwIntegrationDutDataSets[i].CurrentFwName = sDownloadFileName;

//                    int iIndex = 0;

//                    //列出目錄中所有檔案
//                    str = "List all the files in Folder: "  + sDir;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    foreach (string s in list)
//                    {
//                        string[] sFiles = s.Split(' ');
//                        foreach (string st in sFiles)
//                        {
//                            if (st.IndexOf("CGA") >= 0)
//                            {
//                                Invoke(new SetTextCallBackT(SetText), new object[] { st, txtRouterChamberPerformanceFunctionTestInformation });
//                                sFileName[iIndex++] = st;
//                            }
//                        }
//                    }
                    
//                    //確認檔案是否存在
//                    str = "Check if the file exists :" + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    bCheckPoint = false;
//                    foreach (string s in sFileName)
//                    {
//                        if(s == sDownloadFileName)
//                        {
//                            bCheckPoint = true;
//                            break;
//                        }                       
//                    }

//                    if (!bCheckPoint)
//                    { //找不到指定的檔案,等一個小時之後再試
//                        str = "FW File doesn't exist!! " + sDownloadFileName;
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        bRdFwDownloadStatus = false;                      

//                        str = "Wait for one hour and try again!!!";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        i_SetCheckHourTime++;

//                        str = "Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                        return false;
//                    }

//                    //iIndex = 0;
//                    //foreach (string s in sFileName)
//                    //{
//                    //    if (s != null)
//                    //    {
//                    //        string[] sFiles = s.Split(' ');
//                    //        foreach (string st in sFiles)
//                    //        {
//                    //            if (st.ToLower().IndexOf(".bin.cm") < 0 && st.ToLower().IndexOf("test.bin") >= 0)
//                    //            {
//                    //                sFileName[iIndex++] = st;
//                    //                str = "The Download File Exists: " + st;
//                    //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    //            }
//                    //        }
//                    //    }
//                    //}

//                    //iIndex = 0;

//                    ftpclient = null;
//                    ftpclient = new CbtFtpClient(RdPath, "cbtsqa", "jenkins");
//                    str = "Start to Download File: " + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = "";
//                    ftpclient.DownloadFile(RdTempPath, sDir, sDownloadFileName, ref str);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = "Download File Succeed: " + sDownloadFileName;
//                    sInTempFile[i] = sDownloadFileName;
//                    //sFwFiles[i] = sDownloadFileName;                  
                    
//                    //foreach (string st in sFileName)
//                    //{
//                    //    if (st != null)
//                    //    {
//                    //        //sDir = sDir;
//                    //        ftpclient = null;
//                    //        ftpclient = new CbtFtpClient(RdPath, "cbtsqa", "jenkins");
//                    //        str = "Start to Download File: " + st;
//                    //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    //        str = "";
//                    //        ftpclient.DownloadFile(RdTempPath, sDir, st, ref str);
//                    //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    //        str = "Download File Succeed: " + st;
//                    //        sInTempFile[iIndex++] = st;
//                    //    }
//                    //}
//                    ftpclient = null;
//                }
//                catch (Exception ex)
//                {
//                    str = "RD FW Doanload Failed: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    str = "Wait 1 Minutes and try Again !!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    Thread.Sleep(60 *1000);                   
//                    //str = "Start Over!!!";
//                    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    i = -1;
//                }
//            }

//            // RD FW 下載完成  
//            str = "RD FW Doanload Finished!!! Start to Upload to NFS Server!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });            

//            string NfsPath = txtCableSwIntegrationConfigurationNfsServerPath.Text;
//            //上傳 Firmware 檔案到NFS Server 上面去.

//            str = "Check if Local tftp server exists?";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            if (!Directory.Exists(NfsPath))
//            { //連不到NFS的資料夾,等一個小時之後再試
//                str = "Local Tftp Server is unreachable!!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                bRdFwDownloadStatus = false;
//                str = "Wait for one hour and try again!!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                i_SetCheckHourTime++;

//                str = "Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                return false;

//                //bRouterChamberPerformanceTestFTThreadRunning = false;
//                //MessageBox.Show("Tftp Server is unreachable!!!", "Error");
//                ////this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));
//                //threadRouterChamberPerformanceTestFT.Abort();
//                //if (thread_CableSwIntegrationVeriwaveFT == null || thread_CableSwIntegrationVeriwaveFT.ThreadState == System.Threading.ThreadState.Stopped)
//                //    this.Invoke(new showRouterChamberPerformanceTestGUIDelegate(ToggleRouterChamberPerformanceFunctionTestGUI));

                
//                //str = "The Date Folder doesn't exist!! ftp://" + sDir + sDay;
//                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            }

//            str = "Server exist. Copy file to local tftp server: ";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            foreach (string s in sInTempFile)
//            {
//                if (s != null)
//                {
//                    string sFwFile = RdTempPath + "\\" + s;
//                    //string NfsPath = "\\10.6.76.7\tftp root";
//                    //string NfsPath = "";
//                    str = "Copy File to Local NFS server: " + sFwFile;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    string sDestFile = NfsPath + "\\" + s;

//                    //CopyFileToNfsFoler(sFwFile, NfsPath, true);
//                    File.Copy(sFwFile, sDestFile, true);

//                    str = "Copy File Succeed.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                    ////刪除已下載的FW
//                    //if (File.Exists(sFwFile))
//                    //{
//                    //    File.Delete(sFwFile);
//                    //}
//                }
//            }

//            str = "Finished!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            str = "Start to Write FW File Name to each Device Process!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //將要Download的 FW 檔名寫入Device中
//            int snmpVersion = rbtnCableSwIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnCableSwIntegrationConfigurationSnmpV2.Checked ? 2 : 3;

//            for (int i = 0; i < cableSwIntegrationDutDataSets.Length; i++)
//            {
//                string valueResponsed = string.Empty;
//                string expectedValue = string.Empty;
//                string mibType = string.Empty;

//                string sIp = cableSwIntegrationDutDataSets[i].IpAddress;

//                str = "Ping Device First: " + sIp;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                 //Ping Device First
//                /* Ping ip first */
//                //if (!QuickPingRouterChamberPerformanceTest(mtbRouterChamberPerformanceFunctionTestCmIpAddress.Text, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//                if (!QuickPingRouterChamberPerformanceTest(sIp, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//                {
//                    str = " Failed!! ";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    //string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";                    
//                    bRdFwDownloadStatus = false;
//                    str = "Wait for one hour and try again!!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                    i_SetCheckHourTime++;
//                    continue;
//                    //return false;
//                }

//                //Setting Tftp Server IP
//                str = "Setting Tftp Server IP to Device: " + cableSwIntegrationDutDataSets[i].SkuModel + ", " + cableSwIntegrationDutDataSets[i].CurrentFwName;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });                
//                str = "Tftp Server IP: " + mtbCableSwIntegrationConfigurationServerIp.Text;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerOid.Text.Trim(), mibType, mtbCableSwIntegrationConfigurationServerIp.Text.Trim(), snmpVersion);
//                Thread.Sleep(1000);
//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //Setting Tftp Server Address
//                str = "Setting Tftp Server Address Type: " + txtCableSwIntegrationConfigurationTftpServerTypeValue.Text;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerTypeOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerTypeOID.Text.Trim(), mibType, txtCableSwIntegrationConfigurationTftpServerTypeValue.Text.Trim(), snmpVersion);
//                Thread.Sleep(1000);
//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationTftpServerTypeOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                
//                str = "Write Fw File Name to Device: " + cableSwIntegrationDutDataSets[i].SkuModel + ", " + cableSwIntegrationDutDataSets[i].CurrentFwName;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationImageFileOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationImageFileOid.Text.Trim(), mibType, cableSwIntegrationDutDataSets[i].CurrentFwName, snmpVersion);
//                Thread.Sleep(1000);
                
//                // 讀回FW Name
//                valueResponsed = "";
//                mibType = "";

//                cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationImageFileOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
                
//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
                
//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(sFwFileNameOid, "OctetString", sFwFiles[i], snmpVersion);
//                //setStatus = Snmp_CableMibReadWriteTest.SetMibSingleValueByVersion(sOid, mibType, sWriteValue, snmpVersion);              
            
//                //// Set Admin Status to 1 for downloading firmware
//                //valueResponsed = "";
//                //mibType = "";
//                //cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                //cableSwIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), mibType, txtCableSwIntegrationConfigurationAdminStatusValue.Text.Trim(), snmpVersion);
//                //Thread.Sleep(1000);
//                //valueResponsed = "";
//                //mibType = "";
//                //cableSwIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtCableSwIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//                //str = "Mib Type: " + mibType;
//                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//                //str = String.Format("Read Value: {0}", valueResponsed);
//                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });                
//            }

//            str = "Finished!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            //foreach (string s in sInTempFile)
//            //{
//            //    if (s != null)
//            //    {
//            //        string sFwFile = RdTempPath + "\\" + s;
//            //        //刪除已下載的FW
//            //        if (File.Exists(sFwFile))
//            //        {
//            //            File.Delete(sFwFile);
//            //        }
//            //    }
//            //}

//            //Ping Device First
//            /* Ping ip first */
//            //if (!QuickPingRouterChamberPerformanceTest(mtbRouterChamberPerformanceFunctionTestCmIpAddress.Text, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//            //if (!QuickPingRouterChamberPerformanceTest(sIp, Convert.ToInt32(nudRouterChamberPerformanceFunctionTestPingTimeout.Value * 1000)))
//            //{

                
//                //        
//                //        File.WriteAllText(saveFileLog, txtRouterChamberPerformanceFunctionTestInformation.Text);
//                //        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterChamberPerformanceFunctionTestInformation });

//                //        try
//                //        {                   
//                //            xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 2] = "Ping Failed";                  

//                //            /* Write Console Log File as HyperLink in Excel report */
//                //            xls_excelRangeRouterChamberPerformanceTest = xls_excelWorkSheetRouterChamberPerformanceTest.Cells[m_PositionRouterChamberPerformanceTest, 3];
//                //            xls_excelWorkSheetRouterChamberPerformanceTest.Hyperlinks.Add(xls_excelRangeRouterChamberPerformanceTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");               
//                //        }
//                //        catch (Exception ex)
//                //        {
//                //            str = "Write Data to Excel File Exception: " + ex.ToString();
//                //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                //        }
//                //        /* Save end time and close the Excel object */
//                //        xls_excelAppRouterChamberPerformanceTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                //        xls_excelWorkBookRouterChamberPerformanceTest.Save();
//                //        /* Close excel application when finish one job(power level) */
//                //        closeExcelRouterChamberPerformanceTest();
//                //        Thread.Sleep(3000);               
//                //        return false;
//            //}
            
//            // 準備開始Device測試, 下次下載FW時間設定為明天早上5:00
//            i_SetCheckHourTime = 5;
//            i_SetCheckMinuteTime = 00;
//            i_SetCheckSecondTime = 00;
           
//            return true;
//        }
        
//        private bool DownloadFwAndRebootBySnmpRouterChamberPerformanceTest()
//        {






//            return true;
//        }

//        //private bool RebootDeviceRouterChamberPerformanceTest(string sChannel, Comport cComport)
//        //{
//        //    return true;
//        //}

//        private bool CGA2121CmReChanProcess_RouterChamberPerformanceTest(int Freq)
//        {
//            string str = string.Empty;
//            string strCompare = string.Empty;
//            string sGoto_dsCheck = "Moving to Downstream Frequency " + Freq.ToString();
//            //string sGoto_dsCheck = "moving to downstream frequency " + nudCableMibConfigReadWriteFunctionTestGotoChannel.Value.ToString() + "000000 hz";
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds

//            // /* Re-Chan Device */
//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();


//            /* Go to root Directory */
//            if (!CGA2121BackToRootDirectory_CableMibConfigReadWriteTest("cm"))
//            {
//                str = "Change Directory to CM Failed.";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                return false;
//            }

//            str = "ReChan Device with Freq " + Freq.ToString();
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            /* use dogo_ds freq command to re-chan device */
//            //str = "Run goto_ds " + nudCableMibConfigReadWriteFunctionTestGotoChannel.Value.ToString();
//            str = "Run goto_ds 333";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });

//            str = "";

//            stopWatch.Stop();
//            stopWatch.Reset();
//            stopWatch.Restart();

//            long lTimeTemp = stopWatch.ElapsedMilliseconds + 30 * 1000;

//            comPortRouterChamberPerformanceTest.WriteLine("/docsis_ctl/goto_ds " + Freq.ToString());
//            str = comPortRouterChamberPerformanceTest.ReadLine();
//            if (str != "")
//            {
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//            }

//            while (true)
//            {
//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    bRouterChamberPerformanceTestFTThreadRunning = false;
//                    return false;
//                    // Never go here ...
//                }

//                if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value) * 1000)
//                {
//                    b_ReChanOK = false;
//                    return false;
//                }

//                if (str.ToLower().IndexOf(sGoto_dsCheck) != -1)
//                {
//                    b_ReChanOK = true;
//                    break;
//                }

//                Thread.Sleep(100);

//                if (stopWatch.ElapsedMilliseconds > lTimeTemp)
//                {//等待30秒再寫一次, 一直重複寫會有問題
//                    lTimeTemp = stopWatch.ElapsedMilliseconds + 30 * 1000;
//                    comPortRouterChamberPerformanceTest.WriteLine("/docsis_ctl/goto_ds " + Freq.ToString());
//                }
//                str = comPortRouterChamberPerformanceTest.ReadLine();
//                if (str != "")
//                {
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                }
//            }

//            if (!b_ReChanOK) return false;

//            stopWatch.Stop();
//            stopWatch.Reset();
//            stopWatch.Restart();

//            /* Wait for device chan */
//            while (true)
//            {
//                if (bRouterChamberPerformanceTestFTThreadRunning == false)
//                {
//                    bRouterChamberPerformanceTestFTThreadRunning = false;
//                    return false;
//                    // Never go here ...
//                }

//                if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterChamberPerformanceFunctionTestDutRebootTimeout.Value) * 1000)
//                {
//                    break;
//                }

//                str = comPortRouterChamberPerformanceTest.ReadLine();
//                if (str != "")
//                {
//                    strCompare += str;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterChamberPerformanceFunctionTestInformation });
//                }

//                if (strCompare.ToLower().IndexOf("crash") != -1 || strCompare.ToLower().IndexOf("bcm338498") != -1)
//                {
//                    str = "========== System Crash ============";
//                    return false;
//                }
//            }

//            return true;
//        }

//        private bool CGA2121SanityTest_RouterChamberPerformanceTest(int iStartIndex, int iStopIndex)
//        {
//            for (int i = iStartIndex; i <= iStopIndex; i++)
//            {
                

//            }
//            return true;
//        }

//        private void LoadCableSwIntegrationDataGridView()
//        {
//            string sFileName = System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterChamberPerformanceDutsSetting.xml";
//            if(File.Exists(sFileName))
//            {
//                readXmlRouterChamberPerformanceDutsSetting(sFileName);
//                //readXmlRouterChamberPerformanceDutsSetting(System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterChamberPerformanceDutsSetting.xml");
//                //readXmlCableSwIntegrationVeriwaveTestCondition(System.Windows.Forms.Application.StartupPath + "\\testCondition\\CableSwIntegrationVeriwaveTestCondition.xml");
//            }

//            sFileName = System.Windows.Forms.Application.StartupPath + "\\testCondition\\CableSwIntegrationVeriwaveTestCondition.xml";
//            if (File.Exists(sFileName))
//            {
//                readXmlCableSwIntegrationVeriwaveTestCondition(sFileName);
//                //readXmlCableSwIntegrationVeriwaveTestCondition(System.Windows.Forms.Application.StartupPath + "\\testCondition\\CableSwIntegrationVeriwaveTestCondition.xml");
//            }

//        }


//        #endregion





        //private void WebGUIMainFunctionChamberPerformance()
        //{
        //    Thread.Sleep(2000);
        //    CBT_SeleniumApi.GuiScriptParameter ScriptPara = new CBT_SeleniumApi.GuiScriptParameter();
        //    string sExceptionInfo = string.Empty;
        //    string sLoginURL = string.Format("http://{0}:{1}", ls_LoginSettingParametersChamberPerformance.GatewayIP, ls_LoginSettingParametersChamberPerformance.HTTP_Port);
        //    string sCurrentURL = string.Empty;
        //    int DATA_ROW = 15;
        //    bool bTestResult = true;
        //    int j = i_TestScriptIndexChamberPerformance;

        //    ScriptPara.Procedure = st_ReadScriptDataChamberPerformance[j].Procedure;
        //    ScriptPara.Index = st_ReadScriptDataChamberPerformance[j].TestIndex;
        //    ScriptPara.Action = st_ReadScriptDataChamberPerformance[j].Action;
        //    ScriptPara.ActionName = st_ReadScriptDataChamberPerformance[j].ActionName;
        //    ScriptPara.ElementType = st_ReadScriptDataChamberPerformance[j].ElementType;
        //    ScriptPara.ElementXpath = st_ReadScriptDataChamberPerformance[j].ElementXpath;
        //    ScriptPara.ElementXpath = ScriptPara.ElementXpath.Replace('\"', '\'');
        //    ScriptPara.RadioBtnExpectedValueXpath = st_ReadScriptDataChamberPerformance[j].RadioButtonExpectedValueXpath.Replace('\"', '\'');
        //    ScriptPara.WriteValue = st_ReadScriptDataChamberPerformance[j].WriteExpectedValue;
        //    ScriptPara.ExpectedValue = st_ReadScriptDataChamberPerformance[j].WriteExpectedValue;
        //    ScriptPara.URL = sLoginURL + st_ReadScriptDataChamberPerformance[j].WriteExpectedValue;
        //    ScriptPara.TestTimeOut = st_ReadScriptDataChamberPerformance[j].TestTimeOut;
        //    ScriptPara.Note = string.Empty;

        //    ---------------------------------------//
        //    -------------- Go To URL --------------//
        //    ---------------------------------------//
        //    if (j == 0)
        //    {
        //        s_CurrentURLChamberPerformance = sLoginURL;
        //        cs_BrowserChamberPerformance.GoToURL(sLoginURL);
        //        Thread.Sleep(1000);
        //    }
        //    if (ScriptPara.Action.CompareTo("Goto") == 0 && s_CurrentURLChamberPerformance.CompareTo(ScriptPara.URL) != 0)
        //    {
        //        #region Go To URL
        //        s_CurrentURLChamberPerformance = ScriptPara.URL;
        //        Thread.Sleep(1000);
        //        cs_BrowserChamberPerformance.GoToURL(s_CurrentURLChamberPerformance);
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    -------------- Set Value --------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("Set") == 0)
        //    {
        //        #region Set Value
        //        s_InfoStrChamberPerformance = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Set Value", ScriptPara.Index, ScriptPara.ActionName);
        //        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrChamberPerformance, txtChamberPerformanceFunctionTestInformation });

        //        sw_TestTimerChamberPerformance.Stop();
        //        bool checkXPathResult = false;
        //        Stopwatch waitingXpath = new Stopwatch();
        //        waitingXpath.Start();
        //        do
        //        {
        //            checkXPathResult = cs_BrowserChamberPerformance.CheckXPathDisplayed(ScriptPara.ElementXpath);
        //        } while (waitingXpath.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);

        //        if (checkXPathResult == true)
        //        {
        //            Thread.Sleep(2000);
        //        }
        //        else if (checkXPathResult == false)
        //        {
        //            s_InfoStrChamberPerformance = string.Format("...Couldn't find the element!");
        //            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrChamberPerformance, txtChamberPerformanceFunctionTestInformation });
        //        }
        //        waitingXpath.Reset();
        //        sw_TestTimerChamberPerformance.Start();
        //        try
        //        {
        //            if (ScriptPara.WriteValue != "" && st_ReadStepsScriptDataChamberPerformance[i_StepsScriptIndexChamberPerformance].WriteValue[i_CommandDataIndex] != "" && st_ReadStepsScriptDataChamberPerformance[i_StepsScriptIndexChamberPerformance].WriteValue[i_CommandDataIndex] != null)
        //            {
        //                ScriptPara.WriteValue = st_ReadStepsScriptDataChamberPerformance[i_StepsScriptIndexChamberPerformance].WriteValue[i_CommandDataIndex];
        //                i_CommandDataIndex++;
        //            }
        //        }
        //        catch { }
        //        ScriptPara.Note = string.Empty;
        //        try
        //        {
        //            cs_BrowserChamberPerformance.SetWebElementValue(ref ScriptPara);
        //        }
        //        catch
        //        {
        //            ExceptionActionChamberPerformance(s_CurrentURLChamberPerformance, ref ScriptPara);
        //            ScriptPara.Note = "Set Value Error:\n" + ScriptPara.Note;
        //            SwitchToRGconsoleChamberPerformance();
        //            WriteconsoleLogChamberPerformance();
        //            return false;
        //        }
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    -------------- Get Value --------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("Get") == 0)
        //    {
        //        #region Get Value
        //        s_InfoStrChamberPerformance = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Get Value", ScriptPara.Index, ScriptPara.ActionName);
        //        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrChamberPerformance, txtChamberPerformanceFunctionTestInformation });

        //        ScriptPara.Note = string.Empty;
        //        ScriptPara.GetValue = string.Empty;

        //        if (ScriptPara.ElementType.CompareTo("RADIO_BUTTON") == 0)
        //        {
        //            ScriptPara.ElementXpath = st_ReadScriptDataChamberPerformance[j].RadioButtonExpectedValueXpath.Replace('\"', '\'');
        //        }

        //        try
        //        {
        //            bTestResult = cs_BrowserChamberPerformance.GetWebElementValue(ref ScriptPara); // Set Value
        //        }
        //        catch
        //        {
        //            ExceptionActionChamberPerformance(s_CurrentURLChamberPerformance, ref ScriptPara);
        //            ScriptPara.Note = "Execute SubmitButton Error:\n" + ScriptPara.Note;
        //            return false;
        //        }


        //        ---------- Write Test Report ----------//
        //        WriteTestReportChamberPerformance(j, TestResult, ref DATA_ROW, ScriptPara.Note);
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    --------------- ReLogin ---------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("ReLogin") == 0)
        //    {
        //        #region ReLogin
        //        Invoke(new SetTextCallBack(SetText), new object[] { "Log in again...", txtChamberPerformanceFunctionTestInformation });
        //        Re_LoginChamberPerformance();
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    --------------- Alert Login ---------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("AlertLogin") == 0)
        //    {
        //        #region Alert Login
        //        Stopwatch pingTimer = new Stopwatch();
        //        pingTimer.Reset();
        //        pingTimer.Start();
        //        Invoke(new SetTextCallBack(SetText), new object[] { "Login...", txtChamberPerformanceFunctionTestInformation });
        //        while (true)
        //        {
        //            if (pingTimer.ElapsedMilliseconds > (Convert.ToInt32(120000)))
        //            {
        //                break;
        //            }

        //            if (PingClient(ls_LoginSettingParametersChamberPerformance.GatewayIP, 1000))
        //            {
        //                Invoke(new SetTextCallBack(SetText), new object[] { "\tPing Successfully!!", txtChamberPerformanceFunctionTestInformation });
        //                break;
        //            }

        //            Thread.Sleep(1000);
        //            string Info = String.Format(".");
        //            Invoke(new SetTextCallBack(SetText), new object[] { Info, txtChamberPerformanceFunctionTestInformation });
        //        }
        //        Thread.Sleep(3000);
        //        String[] splitLoginInfo = ScriptPara.WriteValue.Split('/');
        //        cs_BrowserChamberPerformance.loginAlertMessage(ScriptPara.URL, splitLoginInfo[0], splitLoginInfo[1]);
        //        Thread.Sleep(3000);
        //        SendKeys.Send(" {Enter}");
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    -------------- File Upload --------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("FileUpload") == 0)
        //    {
        //        #region File Upload
        //        s_InfoStrChamberPerformance = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:File Upload", ScriptPara.Index, ScriptPara.ActionName);
        //        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrChamberPerformance, txtChamberPerformanceFunctionTestInformation });

        //        Thread.Sleep(3000);
        //        ScriptPara.Note = string.Empty;
        //        try
        //        {
        //            cs_BrowserChamberPerformance.fileUploadE8350(ref ScriptPara);
        //        }
        //        catch
        //        {
        //            ExceptionActionChamberPerformance(s_CurrentURLChamberPerformance, ref ScriptPara);
        //            ScriptPara.Note = "Set Value Error:\n" + ScriptPara.Note;
        //            SwitchToRGconsoleChamberPerformance();
        //            WriteconsoleLogChamberPerformance();
        //        }
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    --------------- Wait ---------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("Wait") == 0)
        //    {
        //        #region Wait
        //        ScriptPara.Note = string.Empty;
        //        if (ScriptPara.ElementXpath.CompareTo("") == 0)
        //        {
        //            s_InfoStrChamberPerformance = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Waiting for {2} seconds", ScriptPara.Index, ScriptPara.ActionName, ScriptPara.TestTimeOut);
        //            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrChamberPerformance, txtChamberPerformanceFunctionTestInformation });
        //            sw_TestTimerChamberPerformance.Stop();
        //            Thread.Sleep(Convert.ToInt16(ScriptPara.TestTimeOut) * 1000);
        //            sw_TestTimerChamberPerformance.Start();
        //            SwitchToRGconsoleChamberPerformance();
        //            WriteconsoleLogChamberPerformance();
        //        }
        //        else
        //        {
        //            s_InfoStrChamberPerformance = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Waiting for the XPath, it won't be more than {2} seconds", ScriptPara.Index, ScriptPara.ActionName, ScriptPara.TestTimeOut);
        //            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrChamberPerformance, txtChamberPerformanceFunctionTestInformation });
        //            sw_TestTimerChamberPerformance.Stop();
        //            bool checkXPathResult = false;
        //            Stopwatch waitingXpath = new Stopwatch();
        //            waitingXpath.Start();
        //            do
        //            {
        //                checkXPathResult = cs_BrowserChamberPerformance.CheckXPathDisplayed(ScriptPara.ElementXpath);
        //            } while (waitingXpath.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);
        //            waitingXpath.Reset();
        //            sw_TestTimerChamberPerformance.Start();
        //            SwitchToRGconsoleChamberPerformance();
        //            WriteconsoleLogChamberPerformance();
        //        }
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    --------------- CloseDriver ---------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("CloseDriver") == 0)
        //    {
        //        #region CloseDriver
        //        s_InfoStrChamberPerformance = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Close Driver", ScriptPara.Index, ScriptPara.ActionName);
        //        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrChamberPerformance, txtChamberPerformanceFunctionTestInformation });
        //        cs_BrowserChamberPerformance.Close_WebDriver();
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    --------------- OpenDriver ---------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("OpenDriver") == 0)
        //    {
        //        #region OpenDriver
        //        s_InfoStrChamberPerformance = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Open Driver", ScriptPara.Index, ScriptPara.ActionName);
        //        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrChamberPerformance, txtChamberPerformanceFunctionTestInformation });
        //        TestInitialChamberPerformance();
        //        #endregion
        //    }


        //    WriteconsoleLogChamberPerformance();



        //    string Key1 = "Login-";
        //    string Key2 = "ReLogin";


        //    if (ScriptPara.ActionName.ToLower().IndexOf(Key1.ToLower()) < 0 && ScriptPara.ActionName.ToLower().IndexOf(Key2.ToLower()) < 0)
        //    {
        //        ---------- Write Test Report ----------//
        //        WriteTestReportChamberPerformance(j, TestResult, ref DATA_ROW, ScriptPara.Note);
        //        WriteWebGUITestReportChamberPerformance(bTestResult, ScriptPara);
        //    }

        //    if (ScriptPara.Note != string.Empty)
        //    {
        //        st_ReadFinalScriptDataChamberPerformance[i_FinalScriptIndexChamberPerformance].TestResult = "FAIL";
        //        st_ReadFinalScriptDataChamberPerformance[i_FinalScriptIndexChamberPerformance].Comment = "FAIL in: \n" + st_ReadStepsScriptDataChamberPerformance[i_StepsScriptIndexChamberPerformance].Name;
        //        i_duringStartStopChamberPerformance = Convert.ToInt32(st_ReadFinalScriptDataChamberPerformance[i_FinalScriptIndexChamberPerformance].StopIndex) + 1;
        //    }
        //    bWebGUISingleScriptItemRunning = false;
        //}





        

        
    }
}
