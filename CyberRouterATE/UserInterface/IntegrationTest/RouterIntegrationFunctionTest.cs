///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RouterIntegrationFunctionTest.cs
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
        Excel.Application xls_excelAppRouterIntegrationTest;
        Excel.Workbook xls_excelWorkBookRouterIntegrationTest;
        Excel.Worksheet xls_excelWorkSheetRouterIntegrationTest;
        Excel.Range xls_excelRangeRouterIntegrationTest;

        //int i_BandExcelReportColumnIntegrationTest = 4;
        //int i_ModeExcelReportColumnIntegrationTest = 5;
        //int i_ChannelExcelReportColumnIntegrationTest = 6;
        //int i_BandWidthExcelReportColumnIntegrationTest = 7;
        //int i_SecurityExcelReportColumnIntegrationTest = 8;
        //int i_SecurityKeyExcelReportColumnIntegrationTest = 9;

        int i_BandColumnIntegrationTest = 5;
        int i_ModeColumnIntegrationTest = 6;
        int i_ChannelColumnIntegrationTest = 7;
        int i_BandWidthColumnIntegrationTest = 8;
        int i_SecurityColumnIntegrationTest = 9;
        int i_SecurityKeyColumnIntegrationTest = 10;
        int i_TxResultColumnIntegrationTest = 11;
        int i_RxResultColumnIntegrationTest = 12;
        int i_BiResultColumnIntegrationTest = 13;
        int i_CommentColumnIntegrationTest = 14;
        int i_LogColumnIntegrationTest = 15;


        //===ComPort===        
        Comport comPortRouterIntegrationTest;
        Comport2 comPortRouterIntegrationSwtich;

        //===Thread===
        Thread threadRouterIntegrationTestFT;
        bool bRouterIntegrationTestFTThreadRunning = false;

        //=============
        int m_PositionRouterIntegrationTest = 16;
        int m_LoopRouterIntegrationTest = 1;
        string m_RouterIntegrationTestMainFolder = "";
        string m_RouterIntegrationTestSubFolder = "";

        //RouterDutsSetting[] st_Duts;
        string[] sa_GuiScriptCommandSequenceIntegrationTest;
        string[] sa_GuiScriptValueSequenceIntegrationTest;
        
        //===Selenium===
        CBT_SeleniumApi cs_BrowserIntegrationTest = null;
        string s_CurrentUrlIntegrationTest = string.Empty;
        //string s_CurrentGuiAccessContent = string.Empty;
        

        ////====Email===        
        //string s_EmailSenderID = string.Empty;
        //string s_EmailSenderPassword = string.Empty;
        //string[] s_EmailReceivers = null;

        //===Delegate===
        /* Declare delegate prototype */
        public delegate void showRouterIntegrationTestGUIDelegate();

        #endregion

        /*============================================================================================*/
        /*========================== Controller Event Function Area   ================================*/
        /*============================================================================================*/
        #region

        private void integrationTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sTestItem = Constants.TESTITEM_ROUTER_Integration;

            ToggleToolStripMenuItem(false);
            tabControl_RouterStartPage.Hide();

            integrationTestToolStripMenuItem.Checked = true;
            //            //Hide tpCableIntegrationTemp TabPage
            //            tpCableIntegrationTemp.Parent = null;

            tabControl_RouterIntegration.Show();

            /* Preload settings */
            string xmlFile = System.Windows.Forms.Application.StartupPath + "\\config\\" + Constants.XML_ROUTER_Integration;
            Debug.WriteLine(File.Exists(xmlFile) ? "File exist" : "File is not existing");

            if (!File.Exists(xmlFile))
            {
                writeXmlDefaultRouterIntegrationTest(xmlFile);
            }

            readXmlRouterIntegrationTest(xmlFile);
            tsslMessage.Text = tabControl_RouterIntegration.TabPages[tabControl_RouterIntegration.SelectedIndex].Text + " Control Panel";

            /* Initial Test Condition */
            InitRouterIntegrationDutsSetting();
            InitRouterIntegrationTestCaseSetting();
        }

        private void tabControl_RouterIntegration_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (hasDeleteButton)
            {
                if (btnRouterIntegrationDutsSettingEditSetting.Text == "Cancel")
                {
                    btnRouterIntegrationDutsSettingEditSetting.Text = "Edit";                    
                    dgvRouterIntegrationDutsSettingData.Columns.Remove("Action");
                }

                if (btnRouterIntegrationTestCaseSettingEditSetting.Text == "Cancel")
                {
                    btnRouterIntegrationTestCaseSettingEditSetting.Text = "Edit";                    
                    dgvRouterIntegrationTestCaseSettingData.Columns.Remove("Action"); 
                }

                hasDeleteButton = false;
            }
        }

        private void tabControl_RouterIntegration_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPage.Name == "tpRouterIntegrationDutsSetting")
            {// Get Current comport when Data Sets page is selected
                if (cboxRouterIntegrationDutsSettingComPort.Items.Count == 0)
                {
                    string[] Ports = SerialPort.GetPortNames();
                    cboxRouterIntegrationDutsSettingComPort.SelectedIndex = -1;
                    foreach (string port in Ports)
                    {
                        cboxRouterIntegrationDutsSettingComPort.Items.Add(port);
                        // set index value to 0 if serial device finding
                        cboxRouterIntegrationDutsSettingComPort.SelectedIndex = 0;
                    } //End of foreach
                } //End of if (cboxRouterIntegrationDutsSettingComPort.Items.Count == 0)
            } //End of if (e.TabPage.Name == "tpRouterIntegrationDutsSetting")
        }        
        
        private void btnRouterIntegrationConfigurationLineMessageXmlFileName_Click(object sender, EventArgs e)
        {
            string filename = @"LineGroup.xml";
            string sFilter = "XML file|*.xml|All files (*.*)|*.*";
            string sInitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\";

            txtRouterIntegrationConfigurationLineMessageXmlFileName.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        }

        //private void btnRouterIntegrationConfigurationChariotFolder_Click(object sender, EventArgs e)
        //{
        //    string filename = @"IxChariot.exe";
        //    string sFilter = "Exe file (*.exe)|*.exe|All files (*.*)|*.*";
        //    string sInitialDirectory = @"C:\Program Files\Ixia\IxChariot\";

        //    txtRouterIntegrationConfigurationChariotFolder.Text = LoadFile_Common(filename, sFilter, sInitialDirectory);
        //}

        private void btnRouterIntegrationFunctionTestSaveLog_Click(object sender, EventArgs e)
        {
            SaveLog(txtRouterIntegrationFunctionTestInformation);
        }

        private void btnRouterIntegrationFunctionTestRun_Click(object sender, EventArgs e)
        {            
            /* Prevent double-click from double firing the same action */
            btnRouterIntegrationFunctionTestRun.Enabled = false;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            if (bRouterIntegrationTestFTThreadRunning == false)
            {                
                txtRouterIntegrationFunctionTestInformation.Text = "";               

                if (!ParameterCheckRouterIntegrationTest())
                {
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    btnRouterIntegrationFunctionTestRun.Enabled = true;
                    return;
                }

                if (!InitialRouterIntegrationTest())
                {
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    btnRouterIntegrationFunctionTestRun.Enabled = true;
                    return;
                }

                ///* Create COM port */
                //ReadComportSettingAndInitial();
                //comPortRouterIntegrationTest = new Comport();

                /* Initial Selenium */
                //cs_BrowserIntegrationTest = new CBT_SeleniumApi();

                /* Create sub folder for saving the report data */
                m_RouterIntegrationTestMainFolder = createReportMainFolderRouterIntegration();

                bRouterIntegrationTestFTThreadRunning = true;
                /* Disable all controller */
                ToggleRouterIntegrationFunctionTestController(false);

                btnRouterIntegrationFunctionTestRun.Text = "Stop";
                
                threadRouterIntegrationTestFT =new Thread(new ThreadStart(DoRouterIntegrationFunctionTest));
                threadRouterIntegrationTestFT.Name = "";
                threadRouterIntegrationTestFT.Start();
            }
            else
            {
                tsslMessage.Text = "Function Test Control Panel";
                bRouterIntegrationTestFTThreadRunning = false;
            }                        

            /* Button release */
            Thread.Sleep(3000);
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            btnRouterIntegrationFunctionTestRun.Enabled = true;
        }        

        //                if (comPortRouterIntegrationTest.isOpen() != true)
        //                {
        //                    MessageBox.Show("COM port is not ready!");
        //                    System.Windows.Forms.Cursor.Current = Cursors.Default;
        //                    btnRouterIntegrationFunctionTestRun.Enabled = true;
        //                    return;
        //                }

        //                if (chkRouterIntegrationConfigurationUseSwitch.Checked)
        //                {
        //                    ///* Create COM port */
        //                    ReadComport2SettingAndInitial();

        //                    ///* Create COM port 2 for switch control */
        //                    comPortRouterIntegrationSwtich = new Comport2();

        //                    if (comPortRouterIntegrationSwtich.isOpen() != true)
        //                    {
        //                        MessageBox.Show("COM port 2 is not ready!");
        //                        System.Windows.Forms.Cursor.Current = Cursors.Default;
        //                        btnRouterIntegrationFunctionTestRun.Enabled = true;
        //                        return;
        //                    }                   
        //                }       

        #endregion

        /*============================================================================================*/
        /*=================================== Main Function Area   ===================================*/
        /*============================================================================================*/
        #region
        private void DoRouterIntegrationFunctionTest()
        {
            string str = string.Empty;
            int m_totalTimes = 1;   // Default test time
            m_LoopRouterIntegrationTest = 1; // Reset loop counter
            if (chkRouterIntegrationFunctionTestScheduleOnOff.Checked)
                m_totalTimes = Convert.ToInt32(nudRouterIntegrationFunctionTestTimes.Value);

            /* Start Test Loop */
            do
            {
                if (bRouterIntegrationTestFTThreadRunning == false)
                {
                    bRouterIntegrationTestFTThreadRunning = false;
                    MessageBox.Show("Abort test", "Error");
                    this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                    threadRouterIntegrationTestFT.Abort();
                    // Never go here ...
                }

                m_RouterIntegrationTestMainFolder = CreateMainFolderRouterIntegration();

                if (m_RouterIntegrationTestMainFolder == null || m_RouterIntegrationTestMainFolder == "")
                {
                    str = String.Format("Create Loop {0} Folder Failed. Continue Next Loop.", m_LoopRouterIntegrationTest.ToString());
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                    continue;
                }               

                RouterIntegrationTestMainFunction();

                m_totalTimes--;
                m_LoopRouterIntegrationTest++;
            } while (m_totalTimes != 0);

            bRouterIntegrationTestFTThreadRunning = false;
            this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));      
        }

        private bool RouterIntegrationTestMainFunction()
        {
            //===== Local variable =====

            string str = string.Empty;
            string sExcelFile = string.Empty;
            string sInputFile = string.Empty;
            string sOutputFile = string.Empty;
            string sComment = string.Empty;

            string[] sReport = new string[3];

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();


            //===== Initial variable =====

            /* Initial Value */
            m_PositionRouterIntegrationTest = 16;
            //testConfigRouterIntegrationTest = null;

            str = String.Format("======= Run Main Function ======");
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            if (rbtnRouterIntegrationFunctionTestByDuts.Checked)
            {
                RouterIntegrationTestMainFunctionForDuts();
            }
            else
            {
                RouterIntegrationTestMainFunctionForCase();
            }

            //===== Main Flow =====

            /* Read Duts Setting to run Main function */
            for (int DutsRow = 0; DutsRow < st_DutsSetting.Length; DutsRow++)
            {
                if (bRouterIntegrationTestFTThreadRunning == false)
                {
                    bRouterIntegrationTestFTThreadRunning = false;
                    MessageBox.Show("Abort test", "Error");
                    this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                    threadRouterIntegrationTestFT.Abort();
                    // Never go here ...
                }    

                sInputFile = string.Empty;
                sOutputFile = string.Empty;
                sComment = string.Empty;

                m_PositionRouterIntegrationTest = 16;

                modelInfo.ModelName = st_DutsSetting[DutsRow].ModelName;
                modelInfo.SN = st_DutsSetting[DutsRow].SerialNumber;
                modelInfo.SwVersion = st_DutsSetting[DutsRow].SwVersion;
                modelInfo.HwVersion = st_DutsSetting[DutsRow].HwVersion;
                
                string sChariotTxTstFile = sa_RounterDutsSetting[DutsRow, 13];
                string sChariotRxTstFile = sa_RounterDutsSetting[DutsRow, 14];
                string sChariotBiTstFile = sa_RounterDutsSetting[DutsRow, 15];
                string sRouterIpAddress = st_DutsSetting[DutsRow].IpAddress;
                
                string sGuiScriptExcelFile = st_DutsSetting[DutsRow].GuiScriptExcelFile;
                string sExcelFileName = "IntegrationTest";
                //string subFolder;
                string sLogFile;

                str = String.Format("Run DUT: {0}", modelInfo.ModelName);
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                /* Create Folder */
                str = "Create Folder";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                try
                {
                    m_RouterIntegrationTestSubFolder = createSubFolderRouterIntegration(m_RouterIntegrationTestMainFolder);
                }
                catch (Exception ex)
                {
                    str = "Create Folder Failed. Exception: " + ex.ToString();
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    str = "Create Folder Failed.";
                    sLogFile = m_RouterIntegrationTestMainFolder + "\\" + modelInfo.ModelName + "_HeaderLog_" + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".txt";
                    WriteFailMsgToExcelAndSendEamilIntegrationTest(str, sLogFile);
                    SendLineAlarmIntegrationTest(str);                    
                    continue;
                }
                
                str = "Finished!";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                sLogFile = m_RouterIntegrationTestSubFolder + "\\" + "HeaderLog_" + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".txt";

                #region  marked area
                ///* Read Test Condition xls File */
                //if (!File.Exists(txtRouterIntegrationConfigurationTestConditionExcelFileName.Text))
                //{


                //    m_PositionRouterIntegrationTest = 16;
                //    continue;
                //}


                ///* Read Steps Condition xls File */
                #endregion marked area

                /* Create Folder */
                str = "Read Gui Script";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                /* Check Gui Script File Existence */
                if (!File.Exists(sGuiScriptExcelFile))
                { //Fail. Send Log and Line Alarm
                    str = "Gui Script File doesn't Exist!!";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                                        
                    WriteFailMsgToExcelAndSendEamilIntegrationTest(str, sLogFile);
                    SendLineAlarmIntegrationTest(str);

                    //m_PositionRouterIntegrationTest = 16;
                    continue;
                }

                /* Read Gui Script */
                if (!ConvertExcelToDualArray(sGuiScriptExcelFile, ref sa_RounterGuiScript))
                {//Fail. Send Log and Line Alarm
                    str ="Read Gui Script File Failed!!";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    WriteFailMsgToExcelAndSendEamilIntegrationTest(str, sLogFile);
                    SendLineAlarmIntegrationTest(str);

                    //m_PositionRouterIntegrationTest = 16;
                    continue;                  
                }
                
                str = "Finished!";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                
                string savePath = createReportExcelFileIntegrationTest(m_RouterIntegrationTestSubFolder, modelInfo.ModelName, m_LoopRouterIntegrationTest, sExcelFileName);
                //string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
                //string savePath = createRouterIntegrationTestSaveExcelFile(m_RouterIntegrationTestSubFolder, RouterIntegrationDutDataSets[i].SkuModel, m_LoopRouterIntegrationTest, sExcelFileName);
                /* initial Excel component */
                initialExcelRouterIntegrationTest(savePath);

                try
                {
                    /* Fill Loop, File Name in Excel */
                    xls_excelWorkSheetRouterIntegrationTest.Cells[9, 3] = modelInfo.SN;
                    xls_excelWorkSheetRouterIntegrationTest.Cells[10, 3] = modelInfo.SwVersion;
                    xls_excelWorkSheetRouterIntegrationTest.Cells[11, 3] = modelInfo.HwVersion;
                    xls_excelWorkSheetRouterIntegrationTest.Cells[4, 9] = m_LoopRouterIntegrationTest;
                    xls_excelWorkSheetRouterIntegrationTest.Cells[5, 9] = Path.GetFileName(sGuiScriptExcelFile);
                    xls_excelWorkSheetRouterIntegrationTest.Cells[6, 9] = Path.GetFileName(sChariotTxTstFile);
                    xls_excelWorkSheetRouterIntegrationTest.Cells[7, 9] = Path.GetFileName(sChariotRxTstFile);
                    xls_excelWorkSheetRouterIntegrationTest.Cells[8, 9] = Path.GetFileName(sChariotBiTstFile);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }

                /* Run Main Condition for Excel Data */
                for (int rowCondition = 0; rowCondition < sa_RounterTestCondition.GetLength(0); rowCondition++)
                {
                    if (bRouterIntegrationTestFTThreadRunning == false)
                    {
                        bRouterIntegrationTestFTThreadRunning = false;
                        MessageBox.Show("Abort test", "Error");
                        this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                        threadRouterIntegrationTestFT.Abort();
                        // Never go here ...
                    }

                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

                    if (sa_RounterTestCondition[rowCondition, 0] == "" || sa_RounterTestCondition[rowCondition, 0] == null)
                    {
                        continue;
                    }

                    if (sa_RounterTestCondition[rowCondition, 0].ToLower().Trim() != "v")
                    {
                        continue;
                    }

                    str = String.Format("======= Start Condition Test ======");
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    str = "Model Name : " + modelInfo.ModelName;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    str = "Run Test Item " + rowCondition.ToString();
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    str = "Index is " + sa_RounterTestCondition[rowCondition, 1];
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });


                    str = "Function Name is " + sa_RounterTestCondition[rowCondition, 2];
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    sLogFile = "Log_" + sa_RounterTestCondition[rowCondition, 1] + "_" + sa_RounterTestCondition[rowCondition, 2] + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";

                    try
                    {
                        sLogFile = MakeFilenameValid(sLogFile);
                    }
                    catch (Exception)
                    {
                        sLogFile = "Log_" + sa_RounterTestCondition[rowCondition, 1] + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
                    }

                    sLogFile = m_RouterIntegrationTestSubFolder + "\\" + sLogFile;
                    sReport[2] = sLogFile;
                    sReport[0] = "";
                    sReport[1] = "";

                    st_WifiParameter = new WifiParameter();
                    st_WifiParameter.Band = sa_RounterTestCondition[rowCondition, i_BandColumnIntegrationTest-1];
                    st_WifiParameter.Mode = sa_RounterTestCondition[rowCondition, i_ModeColumnIntegrationTest - 1];
                    st_WifiParameter.Channel = sa_RounterTestCondition[rowCondition, i_ChannelColumnIntegrationTest - 1];
                    st_WifiParameter.BandWidth = sa_RounterTestCondition[rowCondition, i_BandWidthColumnIntegrationTest - 1];

                    if (sa_RounterTestCondition[rowCondition, i_SecurityColumnIntegrationTest - 1] != null && sa_RounterTestCondition[rowCondition, i_SecurityColumnIntegrationTest - 1] != "")
                    {
                        string[] s = sa_RounterTestCondition[rowCondition, i_SecurityColumnIntegrationTest - 1].Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);

                        st_WifiParameter.Security = s[0];
                        if (s.Length > 1)
                            st_WifiParameter.SecurityMode = s[1];
                    }
                    //st_WifiParameter.SecurityIndex = sa_RounterTestCondition[rowCondition, 1];
                    st_WifiParameter.SecurityKey = sa_RounterTestCondition[rowCondition, i_SecurityKeyColumnIntegrationTest-1];

                    try
                    {
                        /* Fill Loop, PowerLevel and Constellation in Excel */
                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = sa_RounterTestCondition[rowCondition, 1];
                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3] = sa_RounterTestCondition[rowCondition, 2];
                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 4] = sa_RounterTestCondition[rowCondition, 3];
                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5] = sa_RounterTestCondition[rowCondition, 4];
                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 6] = sa_RounterTestCondition[rowCondition, 5];
                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7] = sa_RounterTestCondition[rowCondition, 6];
                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 8] = sa_RounterTestCondition[rowCondition, 7];
                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 9] = sa_RounterTestCondition[rowCondition, 8];
                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 10] = sa_RounterTestCondition[rowCondition, 9];
                        //xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 11] = sa_RounterTestCondition[rowCondition, 4];
                        //xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 12] = sa_RounterTestCondition[rowCondition, 4];
                        //xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 13] = sa_RounterTestCondition[rowCondition, 4];
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex);
                    }                    

                    //string sOutputPath = System.Windows.Forms.Application.StartupPath + @"\report\" + Path.GetFileName(m_RouterIntegrationTestSubFolder);
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
                    sOutputFile = m_RouterIntegrationTestSubFolder +"\\" + sOutputHead + "Tx_" + Path.GetFileName(sInputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");

                    //sOutputFile = m_RouterIntegrationTestMainFolder + sOutputHead + "Tx_" + Path.GetFileName(sInputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");


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
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });

                        RouterIntegrationTestReportData(i_TxResultColumnIntegrationTest, sReport);

                        File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
                        //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

                        m_PositionRouterIntegrationTest++;
                        continue;
                            
                    }

                    if(st_WifiParameter.Security.ToLower() != "none")
                    {
                        //if(st_WifiParameter.SecurityMode == null || st_WifiParameter.SecurityMode == "" ||
                        if(st_WifiParameter.SecurityKey == null || st_WifiParameter.SecurityKey == "")
                        {
                            sReport[0] = "Error";
                            sReport[1] = "Wifi Security Parameter couldn't be Empty!";
                            Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
                            Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
                            RouterIntegrationTestReportData(i_TxResultColumnIntegrationTest, sReport);

                            File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
                            //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

                            m_PositionRouterIntegrationTest++;
                            continue;
                        }
                    }

                    /*Initial Setting Sequence */

                    sa_GuiScriptCommandSequenceIntegrationTest[8] = "";
                    sa_GuiScriptCommandSequenceIntegrationTest[9] = "";
                    sa_GuiScriptCommandSequenceIntegrationTest[10] = "";
                    sa_GuiScriptCommandSequenceIntegrationTest[11] = "";
                    sa_GuiScriptCommandSequenceIntegrationTest[12] = "";                    

                    switch (st_WifiParameter.Security.ToLower())
                    {
                        case "none":
                            {
                                sa_GuiScriptCommandSequenceIntegrationTest[7] = "Setting Wifi Security None 5G";
                                sa_GuiScriptValueSequenceIntegrationTest[7] = st_WifiParameter.Security;
                            }
                            break;
                        case "wpa2 personal":
                            {
                                sa_GuiScriptCommandSequenceIntegrationTest[7] = "Setting Wifi Security PSK 5G";
                            }
                            break;
                        case "wpa2 enterprise":
                            {
                                sa_GuiScriptCommandSequenceIntegrationTest[7] = "Setting Wifi Security ENT KEY 5G";
                                sa_GuiScriptCommandSequenceIntegrationTest[8] = "Setting Wifi Security Radius IP.1 5G";
                                sa_GuiScriptCommandSequenceIntegrationTest[9] = "Setting Wifi Security Radius IP.2 5G";
                                sa_GuiScriptCommandSequenceIntegrationTest[10] = "Setting Wifi Security Radius IP.3 5G";
                                sa_GuiScriptCommandSequenceIntegrationTest[11] = "Setting Wifi Security Radius IP.4 5G";
                                sa_GuiScriptCommandSequenceIntegrationTest[12] = "Setting Wifi Security Radius Port 5G";                         
                            }
                            break;
                        default:
                            {                                
                                sReport[0] = "Error";
                                sReport[1] = "Unsupport Security Mode!";
                                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
                                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
                                RouterIntegrationTestReportData(i_TxResultColumnIntegrationTest, sReport);

                                File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
                                //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

                                m_PositionRouterIntegrationTest++;
                                continue;
                            }
                            break;
                    }

                    if(st_WifiParameter.Band.ToLower().IndexOf("2.4") >=0)
                    {//2.4G
                        
                        /* Set SSID */
                        st_WifiParameter.SSID = sa_RounterDutsSetting[DutsRow, 6];

                        for (int i = 0; i < sa_GuiScriptCommandSequenceIntegrationTest.Length; i++)
                        {
                            sa_GuiScriptCommandSequenceIntegrationTest[i] = sa_GuiScriptCommandSequenceIntegrationTest[i].Replace("5G", "2.4G");
                        }                        
                    }
                    else
                    {//5G
                        /* Set SSID */
                        st_WifiParameter.SSID = sa_RounterDutsSetting[DutsRow, 7];

                        for (int i = 0; i < sa_GuiScriptCommandSequenceIntegrationTest.Length; i++)
                        {
                            sa_GuiScriptCommandSequenceIntegrationTest[i] = sa_GuiScriptCommandSequenceIntegrationTest[i].Replace("2.4G", "5G");
                        }                                               
                    }
                    
                    sa_GuiScriptValueSequenceIntegrationTest[1] = st_WifiParameter.SSID;                    
                    sa_GuiScriptValueSequenceIntegrationTest[2] = st_WifiParameter.Mode;                    
                    sa_GuiScriptValueSequenceIntegrationTest[3] = st_WifiParameter.BandWidth;
                    sa_GuiScriptValueSequenceIntegrationTest[4] = st_WifiParameter.Channel;
                    sa_GuiScriptValueSequenceIntegrationTest[5] = "None";
                    sa_GuiScriptValueSequenceIntegrationTest[6] = st_WifiParameter.Security;
                    sa_GuiScriptValueSequenceIntegrationTest[7] = st_WifiParameter.SecurityKey;
                    if(st_WifiParameter.Security.ToLower() == "none")
                    {
                        sa_GuiScriptValueSequenceIntegrationTest[6] = st_WifiParameter.Security;
                        st_WifiParameter.SecurityKey = st_WifiParameter.Security;
                        st_WifiParameter.SecurityMode = st_WifiParameter.Security;
                    }                    
                                        
                    

                    str = "";
                    if (!SettingRouterGuiRouterIntegrationTest(sRouterIpAddress, st_WifiParameter, sa_RounterGuiScript))
                    {
                        str = "Setting GUI Script Fail!! ";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                        sReport[0] = "Fail";
                        sReport[1] = str;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });

                        RouterIntegrationTestReportData(i_TxResultColumnIntegrationTest, sReport);

                        File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
                        //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

                        m_PositionRouterIntegrationTest++;
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
                    if (!RemoteConnectionRouterIntegrationTest(st_WifiParameter, "WifiConnection", ref sReport))
                    {
                        str = "Remote Connection Fail!! ";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                        sReport[0] = "Fail";
                        sReport[1] = str;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });

                        RouterIntegrationTestReportData(i_TxResultColumnIntegrationTest, sReport);

                        File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
                        //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

                        m_PositionRouterIntegrationTest++;
                        continue;
                    }

                    #region Ping Client Card
                    //Ping Client Card
                    string sIp = "";

                    str = "Ping Client Card, IP Address: " + sIp;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    

                    if (!QuickPingIntegrationTest(sIp, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
                    {
                        str = "Ping Failed!! ";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                        sReport[0] = "Fail";
                        sReport[1] = str;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });

                        RouterIntegrationTestReportData(i_TxResultColumnIntegrationTest, sReport);

                        File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
                        //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

                        m_PositionRouterIntegrationTest++;
                        continue;                      
                    }
                    #endregion

                    #region Run Chariot

                    str = "Start to run Chariot";
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    sInputFile = sChariotTxTstFile;
                    str = "Run Chariot Tx tst File: " + sInputFile;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                                       
                    
                    if (!ChariotFunctionRouterIntegrationTest(sInputFile, sOutputFile, ref sResultValue, ref sComment))
                    {
                        sReport[0] = "Fail";
                        sReport[1] = sComment;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
                    }
                    else
                    {
                        sReport[0] = sResultValue;
                        str = "Tx Throughput = " + sResultValue;
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                    }

                    RouterIntegrationTestReportData(i_TxResultColumnIntegrationTest, sReport);

                    sInputFile = sChariotRxTstFile;
                    str = "Run Chariot Rx tst File: " + sInputFile;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    sOutputFile = m_RouterIntegrationTestSubFolder + "\\" + sOutputHead + "Rx_" + Path.GetFileName(sInputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");

                    if (!ChariotFunctionRouterIntegrationTest(sInputFile, sOutputFile, ref sResultValue, ref sComment))
                    {
                        sReport[0] = "Fail";
                        sReport[1] = sComment;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
                    }
                    else
                    {
                        sReport[0] = sResultValue;
                        str = "Rx Throughput = " + sResultValue;
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                    }

                    RouterIntegrationTestReportData(i_RxResultColumnIntegrationTest, sReport);

                    sInputFile = sChariotBiTstFile;
                    str = "Run Chariot Bi tst File: " + sInputFile;
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    sOutputFile = m_RouterIntegrationTestSubFolder + "\\" + sOutputHead + "Bi_" + Path.GetFileName(sInputFile) + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss");

                    if (!ChariotFunctionRouterIntegrationTest(sInputFile, sOutputFile, ref sResultValue, ref sComment))
                    {
                        sReport[0] = "Fail";
                        sReport[1] = sComment;
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
                    }
                    else
                    {
                        sReport[0] = sResultValue;
                        str = "Bi Throughput = " + sResultValue;
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                    }

                    RouterIntegrationTestReportData(i_BiResultColumnIntegrationTest, sReport);
                    //sLogFile = "Log_" + sa_RounterTestCondition[rowCondition, 1] + "_" + sa_RounterTestCondition[rowCondition, 2] + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";

                    //try
                    //{
                    //    sLogFile = MakeFilenameValid(sLogFile);
                    //}
                    //catch (Exception)
                    //{
                    //    sLogFile = "Log_" + sa_RounterTestCondition[rowCondition, 1] + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
                    //}

                    //sLogFile = m_RouterIntegrationTestSubFolder + "\\" + sLogFile;
                    //sReport[2] = sLogFile;

                    //RouterIntegrationTestReportData(11, sReport);
                    
                    #endregion

                    File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);

                    m_PositionRouterIntegrationTest++;
                }// End of Run Final Excel Condition

                /* Save end time and close the Excel object */
                xls_excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                xls_excelWorkBookRouterIntegrationTest.Save();
                /* Close excel application when finish one job(power level) */
                closeExcelRouterIntegrationTest();
                Thread.Sleep(3000);

                if (chkRouterIntegrationConfigurationSendReport.Checked)
                {
                    s_EmailSenderID = txtRouterIntegrationConfigurationReportEmailSenderGmailAccount.Text;
                    s_EmailSenderPassword = txtRouterIntegrationConfigurationReportEmailSenderGmailPassword.Text;
                    s_EmailReceivers = txtRouterIntegrationConfigurationReportEmailSendTo.Text.Split(new string[]{";", ","}, StringSplitOptions.RemoveEmptyEntries);

                    try
                    {
                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Router Chamber Performance ATE Test : " + modelInfo.ModelName + " Test Report", "Test Completed!!!", savePath);
                    }
                    catch (Exception ex)
                    {
                        str = "Send Mail Failed: " + ex.ToString();
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                    }
                }

            }//End of Run Duts Condition

            return true;
        }

        private bool RouterIntegrationTestMainFunctionForDuts()
        {
            string str = string.Empty;

            str = "Run Main Function For Duts.";
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            for (int DutsRow = 0; DutsRow < dgvRouterChamberPerformanceDutsSettingData.RowCount - 1; DutsRow++)
            {
                if (bRouterIntegrationTestFTThreadRunning == false)
                {
                    bRouterIntegrationTestFTThreadRunning = false;
                    MessageBox.Show("Abort test", "Error");
                    this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                    threadRouterIntegrationTestFT.Abort();
                    // Never go here ...
                }

                st_CurrentDut = new RouterDutsSetting();
                st_CurrentDut.index = Convert.ToInt32(dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[0].Value);
                st_CurrentDut.ModelName = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[1].Value.ToString();
                st_CurrentDut.SerialNumber = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[2].Value.ToString();
                st_CurrentDut.SwVersion = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[3].Value.ToString();
                st_CurrentDut.HwVersion = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[4].Value.ToString();
                st_CurrentDut.IpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[5].Value.ToString();
                string sSsid24G = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[6].Value.ToString();
                string sSsid5G = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[7].Value.ToString();
                st_CurrentDut.PcIpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[8].Value.ToString();
                st_CurrentDut.SwitchPort = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[9].Value.ToString();
                st_CurrentDut.ComportNum = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[10].Value.ToString();
                st_CurrentDut.MacAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[11].Value.ToString();
                st_CurrentDut.GuiScriptExcelFile = dgvRouterChamberPerformanceDutsSettingData.Rows[DutsRow].Cells[12].Value.ToString();

                /* Read Current Dut Setting */
                ReadCurrentDutRouterIntegration(DutsRow);

                /* Create Model Folder */
                //CreateDutFolderReadCurrentDutRouterIntegration();

            }
              

            return true;
        }

        private bool RouterIntegrationTestMainFunctionForCase()
        {
            return true;
        }

        /* The function is recorded report data */
        private void RouterIntegrationTestReportData(int iColumn, string[] sReport)
        {
            try
            {
                /* Fill Loop, PowerLevel and Constellation in Excel */
                xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, iColumn] = sReport[0];
                if (sReport[0] != null && sReport[0] != "")
                {
                    if (sReport[0].ToLower() == "fail")
                    {
                        xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, iColumn], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, iColumn]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                }               
                
                if (sReport[1] != null || sReport[1] != "")
                {
                    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, i_CommentColumnIntegrationTest] = sReport[1];
                }

                string sLogFile = sReport[2];
                /* Write Console Log File as HyperLink in Excel report */
                xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, i_LogColumnIntegrationTest];
                xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sLogFile, Type.Missing, "ConsoleLog", "ConsoleLog");                  
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            try
            {
                /* Save Excel */                
                xls_excelWorkBookRouterIntegrationTest.Save();                
            }
            catch (Exception ex)
            {
                //str = "SavWrite Data to Excel File Exception: " + ex.ToString();
                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
            }
        }

        #endregion
        
        //======== Sub-Function of Main
        #region Parameter Check and Data Initialize

        private bool ParameterCheckRouterIntegrationTest()
        {
            /* Config and Check Test Condition */
            SetText("Check Paramters...", txtRouterIntegrationFunctionTestInformation);

            /* Check Duts Setting Content */
            if (dgvRouterIntegrationDutsSettingData.RowCount <= 1)
            {
                MessageBox.Show("Duts Setting Data Area can't be Empty!!");
                return false;
            }

            /* Check Email Parameter */
            if (chkRouterIntegrationConfigurationSendReport.Checked)
            {
                if (txtRouterIntegrationConfigurationReportEmailSendTo.Text == "")
                {
                    MessageBox.Show("Email Send-to list can't be Empty!!");
                    return false;
                }

                if (txtRouterIntegrationConfigurationReportEmailSenderGmailAccount.Text == "")
                {
                    MessageBox.Show("Gmail Account can't be Empty!!");
                    return false;
                }

                if (txtRouterIntegrationConfigurationReportEmailSenderGmailPassword.Text == "")
                {
                    MessageBox.Show("Gmail Password can't be Empty!!");
                    return false;
                }
            }

            /* Check Line Message Parameter */
            if (chkRouterIntegrationConfigurationLineMessage.Checked)
            {
                if (txtRouterIntegrationConfigurationLineMessageXmlFileName.Text == "")
                {
                    MessageBox.Show("Line Group list can't be Empty!!");
                    return false;
                }

                /* Check File Existence */
                if (!File.Exists(txtRouterIntegrationConfigurationLineMessageXmlFileName.Text))
                {
                    MessageBox.Show("Line Group Xml File doesn't exist!!");
                    return false;
                }
            }

            /* Check Swtich Setting */            
            //if (chkRouterIntegrationConfigurationUseSwitch.Checked)
            //{
            //    if (txtRouterIntegrationConfigurationSwitchUsername.Text == "")
            //    {
            //        MessageBox.Show("Switch Login Username can't be Empty!!");
            //        return false;
            //    }

            //    if (txtRouterIntegrationConfigurationSwitchPassword.Text == "")
            //    {
            //        MessageBox.Show("Switch Login Password can't be Empty!!");
            //        return false;
            //    }
            //}

            /* Check Other Setting */
            

            /* Check Radius Server */
            
            /* Check File Existence */
            
            SetText("Finished!", txtRouterIntegrationFunctionTestInformation);

            return true;
        }

        private bool InitialRouterIntegrationTest()
        {
            /* Config and Check Test Condition */
            SetText("Initial Data...", txtRouterIntegrationFunctionTestInformation);

            /* Initial Test Condition */
            SetText("Initial Test Condition Array...", txtRouterIntegrationFunctionTestInformation);
            if (!InitTestConditionArrayRouterIntegrationTest())
                return false;            

            ///* Initial Test Condition */
            //SetText("Initial Steps Condition Array...", txtRouterIntegrationFunctionTestInformation);
            //if (!InitStepsConditionArrayRouterIntegrationTest())
            //    return false;            

            /* Initial Test Condition */
            SetText("Initial Duts Setting Array...", txtRouterIntegrationFunctionTestInformation);
            if (!InitDusSettingArrayRouterIntegrationTest())
                return false;

            SetText("Finished!", txtRouterIntegrationFunctionTestInformation);

            /* Initial Test Condition */
            SetText("Initial Dut GUI Script Command Sequence...", txtRouterIntegrationFunctionTestInformation);
            if (!InitGuiScriptCommandSequenceRouterIntegrationTest())
                return false;

            SetText("Finished!", txtRouterIntegrationFunctionTestInformation);
            
            return true;
        }

        private bool InitTestConditionArrayRouterIntegrationTest()
        {
            //SetText("Initial Test Condition Array...", txtRouterIntegrationFunctionTestInformation);
            string sFileName = "";
            //string[,] saCondition = new string[1, 1];

            if (!ConvertExcelToDualArray(sFileName, ref sa_RounterTestCondition))
            {
                MessageBox.Show("Read Test Condition Failed!!");
                return false;
            }

            return true;
        }

        private bool InitStepsConditionArrayRouterIntegrationTest()
        {
            //SetText("Initial Steps Condition Array...", txtRouterIntegrationFunctionTestInformation);
            string sFileName = "";

            if (!ConvertExcelToDualArray(sFileName, ref sa_RounterStepsCondition))
            {
                MessageBox.Show("Read Steps Condition Failed!!");
                return false;
            }
            return true;
        }

        private bool InitDusSettingArrayRouterIntegrationTest()
        {
            //SetText("Initial Duts Setting Array...", txtRouterIntegrationFunctionTestInformation);

            if (!ConvertDatagridViewDataToDualArray(dgvRouterIntegrationDutsSettingData, ref sa_RounterDutsSetting))
            {
                MessageBox.Show("Read Duts Setting Failed!!");
                return false;
            }

            st_DutsSetting = new RouterDutsSetting[dgvRouterIntegrationDutsSettingData.RowCount - 1];

            for (int i = 0; i < dgvRouterIntegrationDutsSettingData.RowCount - 1; i++)
            {
                st_DutsSetting[i].index = Convert.ToInt32(dgvRouterIntegrationDutsSettingData.Rows[i].Cells[0].Value.ToString());
                st_DutsSetting[i].ModelName = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[1].Value.ToString();
                st_DutsSetting[i].SerialNumber = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[2].Value.ToString();
                st_DutsSetting[i].SwVersion = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[3].Value.ToString();
                st_DutsSetting[i].HwVersion = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[4].Value.ToString();
                st_DutsSetting[i].IpAddress = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[5].Value.ToString();
                //st_DutsSetting[i].SSID = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[6].Value.ToString();
                st_DutsSetting[i].PcIpAddress = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[8].Value.ToString();
                st_DutsSetting[i].SwitchPort = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[9].Value.ToString();
                st_DutsSetting[i].ComportNum = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[10].Value.ToString();
                st_DutsSetting[i].MacAddress = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[11].Value.ToString();
                st_DutsSetting[i].GuiScriptExcelFile = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[12].Value.ToString();
                //st_DutsSetting[i].FwFileName = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[12].Value.ToString();
                //st_DutsSetting[i].SSID = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[13].Value.ToString();
                //st_DutsSetting[i].IpAddress = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[1].Value.ToString();
            }

            return true;
        }

        private bool InitGuiScriptCommandSequenceRouterIntegrationTest()
        {
            SetText("Check if Radius Server IP Valid?", txtRouterIntegrationFunctionTestInformation);
            IPAddress ipaRadiusIP;

            //if (!IPAddress.TryParse(mtbRouterIntegrationConfigurationRadiusServerIpAddress.Text, out ipaRadiusIP))
            //{
            //    SetText("Failed!", txtRouterIntegrationFunctionTestInformation);
            //    return false;
            //}

            //string[] ips = ipaRadiusIP.ToString().Split('.');

       


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

        #region  ReadData(Duts, Test Case)

        private bool ReadCurrentDutRouterIntegration(int Row)
        {
            string str = string.Empty;

            str = String.Format("Read Current DUT Data...");
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            st_CurrentDut = new RouterDutsSetting();
            st_CurrentDut.index = Convert.ToInt32(dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[0].Value);
            st_CurrentDut.ModelName = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[1].Value.ToString();
            st_CurrentDut.SerialNumber = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[2].Value.ToString();
            st_CurrentDut.SwVersion = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[3].Value.ToString();
            st_CurrentDut.HwVersion = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[4].Value.ToString();
            st_CurrentDut.IpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[5].Value.ToString();
            string sSsid24G = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[6].Value.ToString();
            string sSsid5G = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[7].Value.ToString();
            st_CurrentDut.PcIpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[8].Value.ToString();
            st_CurrentDut.SwitchPort = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[9].Value.ToString();
            st_CurrentDut.ComportNum = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[10].Value.ToString();
            st_CurrentDut.MacAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[11].Value.ToString();
            st_CurrentDut.GuiScriptExcelFile = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[12].Value.ToString();

            if(!File.Exists(st_CurrentDut.GuiScriptExcelFile))
            {
                str = String.Format("{0} GUI Script: {1} doesn't Exist!!", st_CurrentDut.ModelName, st_CurrentDut.GuiScriptExcelFile);
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            str = String.Format("Read {0} GUI Script: {1}", st_CurrentDut.ModelName, st_CurrentDut.GuiScriptExcelFile);
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            /* Read GUI Script */
            sa_RounterGuiScript = null;
            GC.Collect();

            /* Read Gui Script */
            if (!ConvertExcelToDualArray(st_CurrentDut.GuiScriptExcelFile, ref sa_RounterGuiScript))
            {
                str = String.Format("Failed!!");
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            str = String.Format("Finished!!");
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            return true;
        }

        private bool ReadCurrentTestCaseRouterIntegration(int Row)
        {
            string str = string.Empty;

            str = String.Format("Read Current Test Case Data...");
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            st_CurrentDut = new RouterDutsSetting();
            st_CurrentDut.index = Convert.ToInt32(dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[0].Value);
            st_CurrentDut.ModelName = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[1].Value.ToString();
            st_CurrentDut.SerialNumber = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[2].Value.ToString();
            st_CurrentDut.SwVersion = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[3].Value.ToString();
            st_CurrentDut.HwVersion = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[4].Value.ToString();
            st_CurrentDut.IpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[5].Value.ToString();
            string sSsid24G = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[6].Value.ToString();
            string sSsid5G = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[7].Value.ToString();
            st_CurrentDut.PcIpAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[8].Value.ToString();
            st_CurrentDut.SwitchPort = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[9].Value.ToString();
            st_CurrentDut.ComportNum = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[10].Value.ToString();
            st_CurrentDut.MacAddress = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[11].Value.ToString();
            st_CurrentDut.GuiScriptExcelFile = dgvRouterChamberPerformanceDutsSettingData.Rows[Row].Cells[12].Value.ToString();

            if (!File.Exists(st_CurrentDut.GuiScriptExcelFile))
            {
                str = String.Format("{0} GUI Script: {1} doesn't Exist!!", st_CurrentDut.ModelName, st_CurrentDut.GuiScriptExcelFile);
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            str = String.Format("Read {0} GUI Script: {1}", st_CurrentDut.ModelName, st_CurrentDut.GuiScriptExcelFile);
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            /* Read GUI Script */
            sa_RounterGuiScript = null;
            GC.Collect();

            /* Read Gui Script */
            if (!ConvertExcelToDualArray(st_CurrentDut.GuiScriptExcelFile, ref sa_RounterGuiScript))
            {
                str = String.Format("Failed!!");
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            str = String.Format("Finished!!");
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            return true;
        }

        #endregion 

        #region Folder related

        private string CreateMainFolderRouterIntegration()
        {
            string str = string.Empty;
            string sPath = string.Empty;

            /* Create Folder */
            str = "Create Main Folder";
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            try
            {
                sPath = System.Windows.Forms.Application.StartupPath + @"\report";
                if (!Directory.Exists(sPath))
                    Directory.CreateDirectory(sPath);

                sPath += @"\RouterTest_" + DateTime.Now.ToString("yyyyMMdd-HHmmss");

                if (!Directory.Exists(sPath))
                    Directory.CreateDirectory(sPath);

                str = "Finished!! Folder Namne: " + sPath; 
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
            }
            catch(Exception ex)
            {
                str = "Failed!! Exception: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
            }

            return sPath;             
        }

        private string createReportMainFolderRouterIntegration()
        {
            string subFolder = DateTime.Now.ToString("yyyyMMdd HH-mm-ss");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

            string Path = System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder;

            return Path;
        }

        private string createSubFolderRouterIntegration(string sMainFolder)
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

        private bool SettingRouterGuiRouterIntegrationTest(string sRouterIP, WifiParameter stWifiParameter, string[,] saRounterGuiScript)
        {
            string str = string.Empty;
            bool bStatus = false;
            bool bFistTime = true;
            bool bGetValue = false;

            if (bRouterIntegrationTestFTThreadRunning == false)
            {
                bRouterIntegrationTestFTThreadRunning = false;
                //MessageBox.Show("Abort test", "Error");
                //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                //threadRouterIntegrationTestFT.Abort();
                return false;
                // Never go here ...
            }

            /* Initial Selenium */
            cs_BrowserIntegrationTest = new CBT_SeleniumApi();
            CBT_SeleniumApi.BrowserType csbtIntegrationTest = CBT_SeleniumApi.BrowserType.Chrome;

            if (!cs_BrowserIntegrationTest.init(csbtIntegrationTest))
            {
                str = "Initial Selenuim Failed.";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            str = "Start Gui Script Setting.";
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            Thread.Sleep(2000);
            LoginSetting cLoginParameter = new LoginSetting();

            CBT_SeleniumApi.GuiScriptParameter GuiScriptParameter = new CBT_SeleniumApi.GuiScriptParameter();
            string sExceptionInfo = string.Empty;
            cLoginParameter.GatewayIP = sRouterIP;
            cLoginParameter.HTTP_Port = "80";
            string sLoginURL = string.Format("http://{0}:{1}", cLoginParameter.GatewayIP, cLoginParameter.HTTP_Port);

            str = "Login URL: " + sLoginURL;
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            string[] GuiCommand = sa_GuiScriptCommandSequenceIntegrationTest;
            string[] GuiWriteValue = sa_GuiScriptValueSequenceIntegrationTest;


            //for(int CommandRow = 0; CommandRow < GuiCommand.Length; CommandRow++)
            for (int CommandRow = 0; CommandRow < GuiCommand.Length; CommandRow++)
            {
                if (bRouterIntegrationTestFTThreadRunning == false)
                {
                    bRouterIntegrationTestFTThreadRunning = false;
                    //MessageBox.Show("Abort test", "Error");
                    //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                    //threadRouterIntegrationTestFT.Abort();
                    return false;
                    // Never go here ...
                }

                s_CurrentGuiAccessContent = string.Empty;
                if (GuiCommand[CommandRow] == null || GuiCommand[CommandRow] == "")
                    continue;

                string sCommand = GuiCommand[CommandRow];
                //string sCommand = "Setting Wifi SSID 2.4G";
                string sWriteValue = GuiWriteValue[CommandRow];

                str = String.Format("Run GUI Script Command: {0}", sCommand);
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                bStatus = false;
                bGetValue = false;

                for (int GuiRow = 0; GuiRow < saRounterGuiScript.GetLength(0); GuiRow++)
                {
                    if (bRouterIntegrationTestFTThreadRunning == false)
                    {
                        bRouterIntegrationTestFTThreadRunning = false;
                        //MessageBox.Show("Abort test", "Error");
                        //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                        //threadRouterIntegrationTestFT.Abort();
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
                        if (!GuiScriptOpenHomePageRouterIntegrationTest(sLoginURL))
                        {
                            cs_BrowserIntegrationTest.Close_WebDriver();
                            cs_BrowserIntegrationTest = null;
                            return false;
                        }

                        bFistTime = false;
                        Thread.Sleep(3000);
                    }

                    str = String.Format("Setting Type: {0}, Index: {1}, Action: {2}, Action Name: {3}, WriteValue: {4}", sCommand, GuiScriptParameter.Index, GuiScriptParameter.Action, GuiScriptParameter.ActionName, GuiScriptParameter.WriteValue);
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    switch (GuiScriptParameter.Action)
                    {
                        case "Set":
                            if (!GuiScriptSetFunctionRouterIntegrationTest(GuiScriptParameter))
                            {
                                cs_BrowserIntegrationTest.Close_WebDriver();
                                cs_BrowserIntegrationTest = null;
                                return false;
                            }
                            break;
                        case "Goto":
                            if (!GuiScriptGotoFunctionRouterIntegrationTest(GuiScriptParameter))
                            {
                                cs_BrowserIntegrationTest.Close_WebDriver();
                                cs_BrowserIntegrationTest = null;
                                return false;
                            }
                            break;
                        case "Get":
                            if (!GuiScriptGetFunctionRouterIntegrationTest(GuiScriptParameter))
                            {
                                cs_BrowserIntegrationTest.Close_WebDriver();
                                cs_BrowserIntegrationTest = null;
                                return false;
                            }                            

                            break;
                        case "Wait":                            
                            if (!GuiScriptWaitFunctionRouterIntegrationTest(GuiScriptParameter))
                            {
                                cs_BrowserIntegrationTest.Close_WebDriver();
                                cs_BrowserIntegrationTest = null;
                                return false;
                            }
                            break;

                        default:
                            {
                                cs_BrowserIntegrationTest.Close_WebDriver();
                                cs_BrowserIntegrationTest = null;
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
                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                    cs_BrowserIntegrationTest.Close_WebDriver();
                    cs_BrowserIntegrationTest = null;
                    return false;
                }

                if (bGetValue)
                {
                    if (s_CurrentGuiAccessContent.Trim() != GuiScriptParameter.WriteValue.Trim())
                    {
                        str = "Get Value Fail! Value is not equal to sWriteValue";
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                        cs_BrowserIntegrationTest.Close_WebDriver();
                        cs_BrowserIntegrationTest = null;
                        return false;
                    }
                }

            } // End of Gui Command for loop

            if (cs_BrowserIntegrationTest != null)
            {
                cs_BrowserIntegrationTest.Close_WebDriver();
                cs_BrowserIntegrationTest = null;
            }


            return true;       
        }

        //---------------------------------------//
        //-------------- Go To URL --------------//
        //---------------------------------------//
        private bool GuiScriptOpenHomePageRouterIntegrationTest(string sLoginURL)
        {
            string str = string.Empty;
            str = string.Format("Open GUI Login URL: {0}", sLoginURL);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            
            s_CurrentUrlIntegrationTest = sLoginURL;
            cs_BrowserIntegrationTest.GoToURL(sLoginURL);
            Thread.Sleep(1000);
            return true;
        }

        //---------------------------------------//
        //-------------- Set Value --------------//
        //---------------------------------------//
        private bool GuiScriptSetFunctionRouterIntegrationTest(CBT_SeleniumApi.GuiScriptParameter GuiScriptParameter)
        {
            string str = string.Empty;
            bool checkXPathResult = false;
            //CBT_SeleniumApi.GuiScriptParameter ScriptParameter = new CBT_SeleniumApi.GuiScriptParameter();

            str = string.Format("Set Parameter to GUI, Index: {0}, Action Name: {1}, Action: Set Value", GuiScriptParameter.Index, GuiScriptParameter.ActionName);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            string sWriteValue = GuiScriptParameter.WriteValue;            

            str = string.Format("Search element: {0}", GuiScriptParameter.ElementXpath);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            /* Find Element, try 1 minutes at most */
            for (int i = 0; i < 30; i++)
            {
                checkXPathResult = cs_BrowserIntegrationTest.CheckXPathDisplayed(GuiScriptParameter.ElementXpath);

                if (checkXPathResult) break;
                else Thread.Sleep(2000);
            }

            if (!checkXPathResult)
            {
                str = string.Format("Failed!");
                Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            str = string.Format("Succeed!");
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            GuiScriptParameter.Note = string.Empty;

            Thread.Sleep(5000);
            /* Start to Set Value */
            str = string.Format("Write Value to DUT: {0}", sWriteValue);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            //if (sWriteValue == "" || sWriteValue == null)
            //{
            //    str = string.Format("Write Value is empty or null.");
            //    Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            //    return false;
            //}

            try
            {
                cs_BrowserIntegrationTest.SetWebElementValue(ref GuiScriptParameter);
            }
            catch (Exception ex)
            {
                ExceptionActionUSBStorage(s_CurrentUrlIntegrationTest, ref GuiScriptParameter);
                GuiScriptParameter.Note = "Set Value Error:\n" + GuiScriptParameter.Note;

                str = "Write Value Exception: " + ex.ToString();
                Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            return true;
        }

        private bool GuiScriptGetFunctionRouterIntegrationTest(CBT_SeleniumApi.GuiScriptParameter GuiScriptParameter)
        {
            string str = string.Empty;
            bool checkXPathResult = false;
            bool bTestResult = true;
            //CBT_SeleniumApi.GuiScriptParameter ScriptParameter = new CBT_SeleniumApi.GuiScriptParameter();

            str = string.Format("Get Parameter to GUI, Index: {0}, Action Name: {1}, Action: Get Value, Expected Value: {2}", GuiScriptParameter.Index, GuiScriptParameter.ActionName, GuiScriptParameter.ExpectedValue);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            string sWriteValue = GuiScriptParameter.WriteValue;

            str = string.Format("Search element: {0}", GuiScriptParameter.ElementXpath);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            /* Find Element, try 1 minutes at most */
            for (int i = 0; i < 30; i++)
            {
                checkXPathResult = cs_BrowserIntegrationTest.CheckXPathDisplayed(GuiScriptParameter.ElementXpath);

                if (checkXPathResult) break;
                else Thread.Sleep(2000);
            }

            if (!checkXPathResult)
            {
                str = string.Format("Failed!");
                Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            //Thread.Sleep(5000);
            str = string.Format("Succeed!");
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            GuiScriptParameter.Note = string.Empty;

            /* Start to Get Value */
            str = string.Format("Start to get value of DUT.");
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            try
            {
                bTestResult = cs_BrowserIntegrationTest.GetWebElementValue(ref GuiScriptParameter);
            }
            catch (Exception ex)
            {
                ExceptionActionUSBStorage(s_CurrentUrlIntegrationTest, ref GuiScriptParameter);
                GuiScriptParameter.Note = "Get Value Error:\n" + GuiScriptParameter.Note;

                str = "Get Value Exception: " + ex.ToString();
                Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            //if (ScriptPara.ElementType.CompareTo("RADIO_BUTTON") == 0)
            //{
            //    ScriptPara.ElementXpath = st_ReadScriptDataUSBStorage[j].RadioButtonExpectedValueXpath.Replace('\"', '\'');
            //}

            //try
            //{
            //    bTestResult = cs_BrowserIntegrationTest..GetWebElementValue(ref ScriptPara); // Set Value
            //}
            //catch
            //{
            //    ExceptionActionUSBStorage(s_CurrentURLUSBStorage, ref ScriptPara);
            //    ScriptPara.Note = "Execute SubmitButton Error:\n" + ScriptPara.Note;
            //    //return false;
            //}

            str = "Read Value is " + GuiScriptParameter.GetValue;
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
            
            s_CurrentGuiAccessContent = GuiScriptParameter.GetValue;
            return true;
        }

        private bool GuiScriptWaitFunctionRouterIntegrationTest(CBT_SeleniumApi.GuiScriptParameter GuiScriptParameter)
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
                if (bRouterIntegrationTestFTThreadRunning == false)
                {
                    bRouterIntegrationTestFTThreadRunning = false;
                    //MessageBox.Show("Abort test", "Error");
                    //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                    //threadRouterIntegrationTestFT.Abort();
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

        private bool GuiScriptGotoFunctionRouterIntegrationTest(CBT_SeleniumApi.GuiScriptParameter GuiScriptParameter)
        {
            string str = string.Empty;            

            str = string.Format("Goto WebPage of GUI, Index: {0}, Action Name: {1}, Action: Set Value", GuiScriptParameter.Index, GuiScriptParameter.ActionName);
            Invoke(new SetTextCallBack(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

            s_CurrentUrlIntegrationTest = GuiScriptParameter.URL;
            Thread.Sleep(1000);
            cs_BrowserIntegrationTest.GoToURL(s_CurrentUrlIntegrationTest);

            return true;
        }

        

        #endregion

        #region Remote Connection
        private bool RemoteConnectionRouterIntegrationTest(WifiParameter stWifiParameter, string sCommand, ref string[] sReport)
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
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
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
            //        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
            //        return false;
            //    }
            //    else
            //    {

            //    }
            //}
           
            str = String.Format("Run Remote Control Function:");
            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
                                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                            }
                            break;                        
                    }                
            }
            catch (Exception ex)
            {
                sReport[0] = "Error";
                sReport[1] = "Setting Remote Control Parameter Error : " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            try
            {
                client = new TcpClient(mtbRouterIntegrationConfigurationRemoteControlServerIpAddress.Text, Convert.ToInt32(nudRouterIntegrationConfigurationRemoteControlServerPort.Value));

                str = "Connecte to server succeed!";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
            }
            catch (Exception ex)
            {
                sReport[0] = "Error";
                sReport[1] = "Connect to Server Failed: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
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
                    if (bRouterIntegrationTestFTThreadRunning == false)
                    {
                        bRouterIntegrationTestFTThreadRunning = false;
                        //MessageBox.Show("Abort test", "Error");
                        //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                        //threadRouterIntegrationTestFT.Abort();
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
                    Invoke(new SetTextCallBackT(SetText), new object[] { receive, txtRouterIntegrationFunctionTestInformation });

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
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
                if (client != null) client.Close();
                return false;
            }          

            /* Check Result */
            int index = str.IndexOf("RESULT");
            if (index < 0)
            {
                sReport[0] = "Error";
                sReport[1] = "Remote Access Error";
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
                return false;
            }

            string[] sResult = str.Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
            if (sResult.Length < 2)
            {
                sReport[0] = "Error";
                sReport[1] = "Remote Access Error";
                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
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
                        //    Invoke(new SetTextCallBackT(SetText), new object[] { sComment, txtRouterIntegrationFunctionTestInformation });
                        //    return false;
                        //}

                        if (sCommandTemp.IndexOf("Release") >= 0)
                        {
                            sReport[0] = sResult[1];
                            Invoke(new SetTextCallBackT(SetText), new object[] { sReport, txtRouterIntegrationFunctionTestInformation });
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
                                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
                            }
                            else
                            {
                                sReport[0] = "FAIL";
                                Invoke(new SetTextCallBackT(SetText), new object[] { sReport[0], txtRouterIntegrationFunctionTestInformation });
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
                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

                        sReport[0] = "Error";
                        sReport[1] = "Unknown Remote Command";
                        Invoke(new SetTextCallBackT(SetText), new object[] { sReport[1], txtRouterIntegrationFunctionTestInformation });
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

        private bool ChariotFunctionRouterIntegrationTest(string sTstFile, string sOutputFile, ref string sResultValue, ref string sComment)
        {
            string str = string.Empty;
            string sInputFile = string.Empty;
            sResultValue = "";

            if (!File.Exists(sTstFile))
            {
                str = "File doesn't exist:" + sTstFile;
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                str = "Skip the item.";
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                sComment = "Tst File doesn't exist!";
                return false;
            }
            
            sInputFile = sTstFile;
            Path_runtst = "\\runtst.exe";
            Path_fmttst = "\\fmttst.exe";            
            
            try
            {
                RunChariotConsole(Path_runtst, Path_fmttst, sInputFile, sOutputFile, txtRouterIntegrationFunctionTestInformation);
                string csvFile = sOutputFile + ".csv";
                sResultValue = ThroughputValue(csvFile);
            }
            catch (Exception ex)
            {
                sResultValue = "";
                str = "Run Chariot Exception: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;
            }                 

            return true;
        }

        #endregion

        #region PingFunction

        private bool QuickPingIntegrationTest(string ip, int iTimeout = 3000)
        {
            /* Ping if the deivce available*/
            bool pingStatus = false;
            //SetText("Ping Device...", txtRouterIntegrationFunctionTestInformation);
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Reset();
            stopWatch.Stop();
            stopWatch.Start();

            while (true)
            {
                if (bRouterIntegrationTestFTThreadRunning == false)
                {
                    bRouterIntegrationTestFTThreadRunning = false;
                    //MessageBox.Show("Abort test", "Error");
                    //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                    //threadRouterIntegrationTestFT.Abort();
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
                //SetText("Ping Failed!!", txtRouterIntegrationFunctionTestInformation);
                //MessageBox.Show("Ping Failed, Abort Test!!", "Error");
                return false;
            }

            //SetText("Ping Succeed!!", txtRouterIntegrationFunctionTestInformation);
            return true;
        }

        #endregion

        private string createReportExcelFileIntegrationTest(string subFolder,string sModelName, int Loop, string FileName)
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

        private bool WriteFailMsgToExcelAndSendEamilIntegrationTest(string sFailMsg, string sLogFile)
        {
            string str = string.Empty;

            try
            {
                xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = sFailMsg;

                /* Write Console Log File as HyperLink in Excel report */
                xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3];
                xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sLogFile, Type.Missing, "ConsoleLog", "ConsoleLog");
            }
            catch (Exception ex)
            {
                str = "Write Data to Excel File Exception: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
            }

            try
            {
                /* Save end time and close the Excel object */
                xls_excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                xls_excelWorkBookRouterIntegrationTest.Save();
                /* Close excel application when finish one job(power level) */
                closeExcelRouterIntegrationTest();
                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                //str = "SavWrite Data to Excel File Exception: " + ex.ToString();
                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                return false;                    
            }

            //string sLogFile = "";
            File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
            //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

            return true;
        }

        private bool SendLineAlarmIntegrationTest(string sFailMsg)
        {
            string str = string.Empty;

            try
            {
                xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = sFailMsg;

                /* Write Console Log File as HyperLink in Excel report */
                xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3];
                //xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
            }
            catch (Exception ex)
            {
                str = "Write Data to Excel File Exception: " + ex.ToString();
                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
            }
            //    /* Save end time and close the Excel object */
            //    xls_excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            //    xls_excelWorkBookRouterIntegrationTest.Save();
            //    /* Close excel application when finish one job(power level) */
            //    closeExcelRouterIntegrationTest();
            //    Thread.Sleep(3000);
            //    //寄出報告
            //    try
            //    {
            //        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Read Excel File Failed", "Read Excel File Failed!!!", savePath);
            //    }
            //    catch (Exception ex)
            //    {
            //        str = "Send Mail Failed: " + ex.ToString();
            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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

        private void ToggleRouterIntegrationFunctionTestGUI()
        {
            ToggleRouterIntegrationFunctionTestController(true);
            Debug.WriteLine("Toggle");
        }

        /* Disable/Enable controller */
        private void ToggleRouterIntegrationFunctionTestController(bool Toggle)
        {
            /* Model Section */            
            btnRouterIntegrationFunctionTestSaveLog.Enabled = Toggle;
            rbtnRouterIntegrationFunctionTestByDuts.Enabled = Toggle;
            rbtnRouterIntegrationFunctionTestByCase.Enabled = Toggle;

            /* Time setting */
            nudRouterIntegrationFunctionTestPingTimeout.Enabled = Toggle;
            nudRouterIntegrationFunctionTestConditionTimeout.Enabled = Toggle;
            nudRouterIntegrationFunctionTestDutRebootTimeout.Enabled = Toggle;

            /* Loop Times */
            chkRouterIntegrationFunctionTestScheduleOnOff.Enabled = Toggle;
            nudRouterIntegrationFunctionTestTimes.Enabled = Toggle;            

            /* Duts Setting Section */
            btnRouterIntegrationDutsSettingAddSetting.Enabled = Toggle;
            btnRouterIntegrationDutsSettingEditSetting.Enabled = Toggle;
            btnRouterIntegrationDutsSettingSaveSetting.Enabled = Toggle;
            btnRouterIntegrationDutsSettingLoadSetting.Enabled = Toggle;
            btnRouterIntegrationDutsSettingMoveUp.Enabled = Toggle;
            btnRouterIntegrationDutsSettingMoveDown.Enabled = Toggle;

            nudRouterIntegrationDutsSettingIndex.Enabled = Toggle;
            txtRouterIntegrationDutsSettingModelName.Enabled = Toggle;
            txtRouterIntegrationDutsSettingSerialNumber.Enabled = Toggle;
            txtRouterIntegrationDutsSettingSwVersion.Enabled = Toggle;
            txtRouterIntegrationDutsSettingHwVersion.Enabled = Toggle;
            mtbRouterIntegrationDutsSettingIpAddress.Enabled = Toggle;
            mtbRouterIntegrationDutsSettingMacAddress.Enabled = Toggle;
            mtbRouterIntegrationDutsSettingPcIpAddress.Enabled = Toggle;
            cboxRouterIntegrationDutsSettingComPort.Enabled = Toggle; 
            nudRouterIntegrationDutsSettingwitchPort.Enabled = Toggle;
            txtRouterIntegrationDutsSetting24gSsid.Enabled = Toggle;
            txtRouterIntegrationDutsSettingGuiScripExcelFileName.Enabled = Toggle;
            txtRouterIntegrationDutsSetting5gSsid.Enabled = Toggle;         

            /* Test Case Setting Section */
            btnRouterIntegrationTestCaseSettingAddSetting.Enabled = Toggle;
            btnRouterIntegrationTestCaseSettingEditSetting.Enabled = Toggle;
            btnRouterIntegrationTestCaseSettingSaveSetting.Enabled = Toggle;
            btnRouterIntegrationTestCaseSettingLoadSetting.Enabled = Toggle;
            btnRouterIntegrationTestCaseSettingMoveUp.Enabled = Toggle;
            btnRouterIntegrationTestCaseSettingMoveDown.Enabled = Toggle;            

            nudRouterIntegrationTestCaseSettingIndex.Enabled = Toggle;
            txtRouterIntegrationTestCaseSettingItemName.Enabled = Toggle;
            cboxRouterIntegrationTestCaseSettingItemFunctionType.Enabled = Toggle;
            chkRouterIntegrationTestCaseSettingDutResetToDefaultAfterTest.Enabled = Toggle;
            chkRouterIntegrationTestCaseSettingDutRebootAfterTest.Enabled = Toggle;
            nudRouterIntegrationTestCaseSettingWaitTimeAfterTest.Enabled = Toggle;
            txtRouterIntegrationTestCaseSettingFinalExcelFile.Enabled = Toggle;
            txtRouterIntegrationTestCaseSettingStepExcelFile.Enabled = Toggle;
            btnRouterIntegrationTestCaseSettingFinalExcelFile.Enabled = Toggle;
            btnRouterIntegrationTestCaseSettingStepExcelFile.Enabled = Toggle;           

            /* Email Setting */
            chkRouterIntegrationConfigurationSendReport.Enabled = Toggle;
            txtRouterIntegrationConfigurationReportEmailSendTo.Enabled = Toggle;
            txtRouterIntegrationConfigurationReportEmailSenderGmailAccount.Enabled = Toggle;
            txtRouterIntegrationConfigurationReportEmailSenderGmailPassword.Enabled = Toggle;
            
            /* Line Message Setting */
            chkRouterIntegrationConfigurationLineMessage.Enabled = Toggle;
            txtRouterIntegrationConfigurationLineMessageXmlFileName.Enabled = Toggle;
            btnRouterIntegrationConfigurationLineMessageXmlFileName.Enabled = Toggle; 

            /* Remote Control Setting */
            chkRouterIntegrationRemoteControl.Enabled = Toggle;
            mtbRouterIntegrationConfigurationRemoteControlServerIpAddress.Enabled = Toggle;
            nudRouterIntegrationConfigurationRemoteControlServerPort.Enabled = Toggle;
            chkRouterIntegrationClientCard.Enabled = Toggle;
            mtbRouterIntegrationConfigurationClientCardIpAddress.Enabled = Toggle;

            /* Swtich Setting */
            //chkRouterIntegrationConfigurationUseSwitch.Enabled = Toggle;
            //txtRouterIntegrationConfigurationSwitchUsername.Enabled = Toggle;
            //txtRouterIntegrationConfigurationSwitchPassword.Enabled = Toggle;           

            btnRouterIntegrationFunctionTestSaveLog.Enabled = Toggle;
            btnRouterIntegrationFunctionTestRun.Text = Toggle ? "Run" : "Stop";
            Debug.WriteLine(btnRouterIntegrationFunctionTestRun.Text);

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

            if (bRouterIntegrationTestFTThreadRunning == false)
            {
                /* Save excel data */
                if (xls_excelWorkBookRouterIntegrationTest != null)
                {
                    try
                    {
                        /* Save end time and close the Excel object */
                        xls_excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                        xls_excelWorkBookRouterIntegrationTest.Save();
                        closeExcelRouterIntegrationTest();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Save Excel Error: " + ex.ToString());
                    }
                }

                if (cs_BrowserIntegrationTest != null)
                    cs_BrowserIntegrationTest.Close_WebDriver();

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
        private void initialExcelRouterIntegrationTest(string savePath)
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
                            xls_excelAppRouterIntegrationTest= new Excel.Application();
                        }
                        else /* Use exist process */
                        {
                            object obj = Marshal.GetActiveObject("Excel.Application");
                            xls_excelAppRouterIntegrationTest = obj as Excel.Application;                 
                        }

                        if(xls_excelAppRouterIntegrationTest !=null)
                            break;

                        Thread.Sleep(3000*(i+1));
                    }
#else
            xls_excelAppRouterIntegrationTest = new Excel.Application();
#endif
            if (xls_excelAppRouterIntegrationTest == null)
            {
                bRouterIntegrationTestFTThreadRunning = false;
                MessageBox.Show("Excel process could not be forked !!!", "Error");
                //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
                //threadRouterIntegrationTestFT.Abort();
                return;
                //Never go here
            }

            //Maxmimize excel windows
            xls_excelAppRouterIntegrationTest.WindowState = Excel.XlWindowState.xlMaximized;

            /* Set Excel visible */
            xls_excelAppRouterIntegrationTest.Visible = true;

            /* do not show alert */
            xls_excelAppRouterIntegrationTest.DisplayAlerts = false;

            xls_excelAppRouterIntegrationTest.UserControl = true;
            xls_excelAppRouterIntegrationTest.Interactive = false;


            /* Set font and font size attributes */
            xls_excelAppRouterIntegrationTest.StandardFont = "Arial";
            xls_excelAppRouterIntegrationTest.StandardFontSize = 10;

            xls_excelWorkBookRouterIntegrationTest = xls_excelAppRouterIntegrationTest.Workbooks.Add(misValue); /* This method is used to open an Excel workbook by passing the file path as a parameter to this method. */

            xls_excelWorkSheetRouterIntegrationTest = (Excel.Worksheet)xls_excelWorkBookRouterIntegrationTest.Sheets[1]; /* By default, every workbook created has three worksheets */
            xls_excelWorkSheetRouterIntegrationTest.Name = "Results";
            createTitleExcelRouterIntegrationTest();

            try
            {
                xls_excelWorkBookRouterIntegrationTest.SaveAs(savePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp);
            }
        }

        private void createTitleExcelRouterIntegrationTest()
        {
            xls_excelWorkSheetRouterIntegrationTest.Cells[2, 2] = "CyberATE " + modelInfo.ModelName + " Router Chamber Performance Test Report";
            xls_excelWorkSheetRouterIntegrationTest.Cells[4, 2] = "Test";
            xls_excelWorkSheetRouterIntegrationTest.Cells[5, 2] = "Station";
            xls_excelWorkSheetRouterIntegrationTest.Cells[6, 2] = "Start Time";
            xls_excelWorkSheetRouterIntegrationTest.Cells[7, 2] = "End Time";
            xls_excelWorkSheetRouterIntegrationTest.Cells[8, 2] = "Model";
            xls_excelWorkSheetRouterIntegrationTest.Cells[9, 2] = "Serial";
            xls_excelWorkSheetRouterIntegrationTest.Cells[10, 2] = "SW Version";
            xls_excelWorkSheetRouterIntegrationTest.Cells[11, 2] = "HW Version";

            xls_excelWorkSheetRouterIntegrationTest.Cells[4, 3] = modelInfo.ModelName + " Router Chamber Performance Test";
            xls_excelWorkSheetRouterIntegrationTest.Cells[5, 3] = "CyberTAN Router-" + modelInfo.ModelName;

            xls_excelWorkSheetRouterIntegrationTest.Cells[6, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            xls_excelWorkSheetRouterIntegrationTest.Cells[8, 3] = modelInfo.ModelName;
            xls_excelWorkSheetRouterIntegrationTest.Cells[9, 3] = modelInfo.SN;
            xls_excelWorkSheetRouterIntegrationTest.Cells[10, 3] = modelInfo.SwVersion;
            xls_excelWorkSheetRouterIntegrationTest.Cells[11, 3] = modelInfo.HwVersion;

            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[6, 3], xls_excelWorkSheetRouterIntegrationTest.Cells[11, 3]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            /* Set cells width */
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 1], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 1]].ColumnWidth = 2;
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 2], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 2]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 3], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 3]].ColumnWidth = 15;
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 4], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 4]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 5]].ColumnWidth = 15;
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 6], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 6]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 7], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 7]].ColumnWidth = 15;
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 8], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 8]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 9], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 9]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 10], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 10]].ColumnWidth = 15; /* column B */
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[2, 11], xls_excelWorkSheetRouterIntegrationTest.Cells[10, 11]].ColumnWidth = 15; /* column B */

            xls_excelAppRouterIntegrationTest.Cells[3, 8] = "Option";
            xls_excelAppRouterIntegrationTest.Cells[3, 9] = "Value";

            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelAppRouterIntegrationTest.Cells[3, 8], xls_excelAppRouterIntegrationTest.Cells[8, 9]].Borders.LineStyle = 1;

            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelAppRouterIntegrationTest.Cells[3, 8], xls_excelAppRouterIntegrationTest.Cells[3, 9]].Font.Underline = true;
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelAppRouterIntegrationTest.Cells[3, 8], xls_excelAppRouterIntegrationTest.Cells[3, 9]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelAppRouterIntegrationTest.Cells[3, 8], xls_excelAppRouterIntegrationTest.Cells[3, 9]].Font.FontStyle = "Bold";

            xls_excelWorkSheetRouterIntegrationTest.Cells[4, 8] = "Loop";
            xls_excelWorkSheetRouterIntegrationTest.Cells[5, 8] = "Gui Script File";
            xls_excelWorkSheetRouterIntegrationTest.Cells[6, 8] = "Tx Tst";
            xls_excelWorkSheetRouterIntegrationTest.Cells[7, 8] = "Rx Tst";
            xls_excelWorkSheetRouterIntegrationTest.Cells[8, 8] = "Bi Tst";

            /* Fill the Title */
            int rowCount = 15;
            //xls_excelAppRouterIntegrationTest.Cells[rowCount, 2] = "Index";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 2] = "Index";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 3] = "Function Name";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 4] = "Name";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 5] = "Band";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 6] = "Mode";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 7] = "Channel";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 8] = "BandWidth";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 9] = "Security";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 10] = "Security Key";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 11] = "Tx(Mbps)";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 12] = "Rx(Mbps)";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 13] = "Bi(Mbps)";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, i_CommentColumnIntegrationTest] = "Comment";
            xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, i_LogColumnIntegrationTest] = "Log File";
            //xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 10] = "Log File";
            //xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 11] = "Test Result";
            //xls_excelWorkSheetRouterIntegrationTest.Cells[rowCount, 12] = "Comment";

            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[14, 2], xls_excelWorkSheetRouterIntegrationTest.Cells[14, 10]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }      

        private void saveExcelRouterIntegrationTest(string savePath)
        {
            try
            {
                xls_excelWorkBookRouterIntegrationTest.SaveAs(savePath, misValue, misValue, misValue,
                    misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue,
                    misValue, misValue, misValue);
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp);
            }
        }       

        private void closeExcelRouterIntegrationTest()
        {
            /* Turn on interactive mode */
            xls_excelAppRouterIntegrationTest.Interactive = true;
            xls_excelWorkBookRouterIntegrationTest.Close();
            xls_excelAppRouterIntegrationTest.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xls_excelWorkSheetRouterIntegrationTest);
            xls_excelWorkSheetRouterIntegrationTest = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xls_excelWorkBookRouterIntegrationTest);
            xls_excelWorkBookRouterIntegrationTest = null;

            if (xls_excelRangeRouterIntegrationTest != null)
            {
                releaseObject_RouterIntegrationTest(xls_excelRangeRouterIntegrationTest);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xls_excelRangeRouterIntegrationTest);
                xls_excelRangeRouterIntegrationTest = null;
            }
            releaseObject_RouterIntegrationTest(xls_excelAppRouterIntegrationTest);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xls_excelAppRouterIntegrationTest);
            xls_excelAppRouterIntegrationTest = null;

            GC.Collect();
        }

        private void releaseObject_RouterIntegrationTest(object obj)
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
        public void writeXmlDefaultRouterIntegrationTest(string FileName)
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

            //// Model section
            //writer.WriteStartElement("Model");
            //writer.WriteElementString("Name", "CGA2121");
            //writer.WriteElementString("SN", "0123456789");
            //writer.WriteElementString("SWVer", "Diag. 1.0");
            //writer.WriteElementString("HWVer", "ES 1.0");
            //writer.WriteEndElement();

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

            // Remote Control Setting Section
            writer.WriteStartElement("RemoteControlSetting");
            writer.WriteElementString("EnableRemoteControl", "Y");
            writer.WriteElementString("RemoteServerIp", "192.168.0.12");
            writer.WriteElementString("RemoteServerPort", "7890");
            writer.WriteElementString("UseClientCard", "Y");
            writer.WriteElementString("ClientCardIp", "192.168.1.100");
            writer.WriteEndElement();

            //// Switch Section
            //writer.WriteStartElement("SwitchSetting");
            //writer.WriteElementString("UseSwitch", "N");
            //writer.WriteElementString("Username", "cisco");
            //writer.WriteElementString("Password", "sqa11063");
            //writer.WriteEndElement();

            //// Other Setting Section
            //writer.WriteStartElement("OtherSetting");
            //writer.WriteElementString("RemoteServerIp", "192.168.0.12");
            //writer.WriteElementString("RemoteServerPort", "7890");
            //writer.WriteElementString("ClientCardIp", "192.168.1.100");
            //writer.WriteElementString("ChariotPath", @"C:\Program Files\Ixia\IxChariot\IxChariot.exe");
            //writer.WriteElementString("TestConditionFileName", "");
            //writer.WriteEndElement();

            //// Radius Server Section
            //writer.WriteStartElement("RadiusServerSetting");
            //writer.WriteElementString("RadiusServerIp", "192.168.1.100");
            //writer.WriteElementString("RadiusServerPort", "1234");            
            //writer.WriteEndElement();

            writer.WriteEndElement();
            // End of write Configuration

            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public void writeXmlRouterIntegrationTest(string FileName)
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

            //// Model section
            //writer.WriteStartElement("Model");
            //writer.WriteElementString("Name", txtRouterIntegrationFunctionTestName.Text);
            //writer.WriteElementString("SN", txtRouterIntegrationFunctionTestSerialNumber.Text);
            //writer.WriteElementString("SWVer", txtRouterIntegrationFunctionTestSwVersion.Text);
            //writer.WriteElementString("HWVer", txtRouterIntegrationFunctionTestHwVersion.Text);
            //writer.WriteEndElement();

            /* Time Setting*/
            writer.WriteStartElement("TimeSetting");
            writer.WriteElementString("PingTimeout", nudRouterIntegrationFunctionTestPingTimeout.Value.ToString());
            writer.WriteElementString("ConditionTimeout", nudRouterIntegrationFunctionTestConditionTimeout.Value.ToString());
            writer.WriteElementString("DutRebootTimeout", nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString());
            writer.WriteEndElement();

            /* Loop Times */
            writer.WriteStartElement("LoopSetting");
            writer.WriteElementString("ScheduleOnOff", chkRouterIntegrationFunctionTestScheduleOnOff.Checked ? "On" : "Off");
            writer.WriteElementString("ScheduleTimes", nudRouterIntegrationFunctionTestTimes.Value.ToString());
            writer.WriteEndElement();

            writer.WriteEndElement();
            // End of write Function Test

            ///
            /// Write Configuration settings
            /// 
            writer.WriteStartElement("Configuration");

            //Email Section
            writer.WriteStartElement("EmailSetting");
            writer.WriteElementString("SendReport", chkRouterIntegrationConfigurationSendReport.Checked? "Y":"N");
            writer.WriteElementString("ReportSenderAccount", txtRouterIntegrationConfigurationReportEmailSenderGmailAccount.Text);
            writer.WriteElementString("ReportSenderPassword", txtRouterIntegrationConfigurationReportEmailSenderGmailPassword.Text);
            writer.WriteElementString("ReportSendTo", txtRouterIntegrationConfigurationReportEmailSendTo.Text);
            writer.WriteEndElement();

            // Line Section
            writer.WriteStartElement("LineSetting");
            writer.WriteElementString("LineMessage", chkRouterIntegrationConfigurationLineMessage.Checked ? "Y" : "N");
            writer.WriteElementString("LineGroupXmlFile", txtRouterIntegrationConfigurationLineMessageXmlFileName.Text);
            writer.WriteEndElement();

            // Remote Control Setting Section
            writer.WriteStartElement("RemoteControlSetting");
            writer.WriteElementString("EnableRemoteControl", chkRouterIntegrationRemoteControl.Checked? "Y":"N");
            writer.WriteElementString("RemoteServerIp", mtbRouterIntegrationConfigurationRemoteControlServerIpAddress.Text);
            writer.WriteElementString("RemoteServerPort", nudRouterIntegrationConfigurationRemoteControlServerPort.Value.ToString());
            writer.WriteElementString("UseClientCard", chkRouterIntegrationClientCard.Checked ? "Y" : "N");
            writer.WriteElementString("ClientCardIp", mtbRouterIntegrationConfigurationClientCardIpAddress.Text);
            writer.WriteEndElement();

            ////Switch Section
            //writer.WriteStartElement("SwitchSetting");
            //writer.WriteElementString("UseSwitch", chkRouterIntegrationConfigurationUseSwitch.Checked ? "Y" : "N");
            //writer.WriteElementString("Username", txtRouterIntegrationConfigurationSwitchUsername.Text);
            //writer.WriteElementString("Password", txtRouterIntegrationConfigurationSwitchPassword.Text);
            //writer.WriteEndElement();                    

            ////Other Setting
            //writer.WriteStartElement("OtherSetting");
            
            //writer.WriteElementString("ChariotPath", txtRouterIntegrationConfigurationChariotFolder.Text);
            //writer.WriteElementString("TestConditionFileName", txtRouterIntegrationConfigurationTestConditionExcelFileName.Text);
            //writer.WriteEndElement();

            //// Radius Server Section
            //writer.WriteStartElement("RadiusServerSetting");
            //writer.WriteElementString("RadiusServerIp", mtbRouterIntegrationConfigurationRadiusServerIpAddress.Text);
            //writer.WriteElementString("RadiusServerPort", nudRouterIntegrationConfigurationRadiusServerPort.Value.ToString());
            //writer.WriteEndElement();

            writer.WriteEndElement();
            // End of write Configuration          

            writer.WriteEndDocument();

            writer.Flush();
            writer.Close();
        }

        public bool readXmlRouterIntegrationTest(string FileName)
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
            //// Model
            //XmlNode nodeFunctionTestModel = doc.SelectSingleNode("/RouterATE/FunctionTest/Model");
            //try
            //{
            //    string Name = nodeFunctionTestModel.SelectSingleNode("Name").InnerText;
            //    string SN = nodeFunctionTestModel.SelectSingleNode("SN").InnerText;
            //    string SWVer = nodeFunctionTestModel.SelectSingleNode("SWVer").InnerText;
            //    string HWVer = nodeFunctionTestModel.SelectSingleNode("HWVer").InnerText;

            //    Debug.WriteLine("Name: " + Name);
            //    Debug.WriteLine("SN: " + SN);
            //    Debug.WriteLine("SWVer: " + SWVer);
            //    Debug.WriteLine("HWVer: " + HWVer);

            //    txtRouterIntegrationFunctionTestName.Text = Name;
            //    txtRouterIntegrationFunctionTestSerialNumber.Text = SN;
            //    txtRouterIntegrationFunctionTestSwVersion.Text = SWVer;
            //    txtRouterIntegrationFunctionTestHwVersion.Text = HWVer;
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine("/RouterATE/FunctionTest/Model " + ex);
            //}

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

                nudRouterIntegrationFunctionTestPingTimeout.Value = Convert.ToDecimal(PingTimeout);
                nudRouterIntegrationFunctionTestConditionTimeout.Value = Convert.ToDecimal(ConditionTimeout);
                nudRouterIntegrationFunctionTestDutRebootTimeout.Value = Convert.ToDecimal(DutRebootTimeout);
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

                chkRouterIntegrationFunctionTestScheduleOnOff.Checked = (ScheduleOnOff == "On" ? true : false);
                nudRouterIntegrationFunctionTestTimes.Value = Convert.ToDecimal(ScheduleTimes);
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
                    chkRouterIntegrationConfigurationSendReport.Checked = true;
                }
                else
                {
                    chkRouterIntegrationConfigurationSendReport.Checked = false;
                }

                txtRouterIntegrationConfigurationReportEmailSenderGmailAccount.Text = ReportSenderAccount;
                txtRouterIntegrationConfigurationReportEmailSenderGmailPassword.Text = ReportSenderPassword;
                txtRouterIntegrationConfigurationReportEmailSendTo.Text = ReportSendTo;
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
                    chkRouterIntegrationConfigurationLineMessage.Checked = true;
                }
                else
                {
                    chkRouterIntegrationConfigurationLineMessage.Checked = false;
                }
                
                txtRouterIntegrationConfigurationLineMessageXmlFileName.Text = LineGroupXmlFile;
                
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATE/Configuration/LineSetting " + ex);
            }

            // Remote Control Setting 
            XmlNode nodeConfigurationRemoteControlSetting = doc.SelectSingleNode("/RouterATE/Configuration/RemoteControlSetting");
            try
            {
                string EnableRemoteControl = nodeConfigurationRemoteControlSetting.SelectSingleNode("EnableRemoteControl").InnerText;
                string RemoteServerIp = nodeConfigurationRemoteControlSetting.SelectSingleNode("RemoteServerIp").InnerText;
                string RemoteServerPort = nodeConfigurationRemoteControlSetting.SelectSingleNode("RemoteServerPort").InnerText;
                string UseClientCard = nodeConfigurationRemoteControlSetting.SelectSingleNode("EnableRemoteControl").InnerText;
                string ClientCardIp = nodeConfigurationRemoteControlSetting.SelectSingleNode("ClientCardIp").InnerText;


                Debug.WriteLine("EnableRemoteControl: " + EnableRemoteControl);
                Debug.WriteLine("RemoteServerIp: " + RemoteServerIp);
                Debug.WriteLine("RemoteServerPort: " + RemoteServerPort);
                Debug.WriteLine("UseClientCard: " + UseClientCard);
                Debug.WriteLine("ClientCardIp: " + ClientCardIp);

                mtbRouterIntegrationConfigurationRemoteControlServerIpAddress.Text = RemoteServerIp;
                nudRouterIntegrationConfigurationRemoteControlServerPort.Text = RemoteServerPort;
                mtbRouterIntegrationConfigurationClientCardIpAddress.Text = ClientCardIp;
                chkRouterIntegrationRemoteControl.Checked = (EnableRemoteControl == "Y"? true:false);
                chkRouterIntegrationClientCard.Checked = (UseClientCard == "Y"? true:false);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/RouterATE/Configuration/OtherSetting " + ex);
            }

            //// Switch Setting 
            //XmlNode nodeConfigurationSwitchSetting = doc.SelectSingleNode("/RouterATE/Configuration/SwitchSetting");
            //try
            //{
            //    string UseSwitch = nodeConfigurationSwitchSetting.SelectSingleNode("UseSwitch").InnerText;
            //    string Username = nodeConfigurationSwitchSetting.SelectSingleNode("Username").InnerText;
            //    string Password = nodeConfigurationSwitchSetting.SelectSingleNode("Password").InnerText;

            //    Debug.WriteLine("UseSwitch: " + UseSwitch);
            //    Debug.WriteLine("Username: " + Username);
            //    Debug.WriteLine("Password: " + Password);

            //    chkRouterIntegrationConfigurationUseSwitch.Checked = (UseSwitch == "Y" ? true : false);
            //    txtRouterIntegrationConfigurationSwitchUsername.Text = Username;
            //    txtRouterIntegrationConfigurationSwitchPassword.Text = Password;
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine("/RouterATE/Configuration/SwitchSetting " + ex);
            //}

            //// Other Setting 
            //XmlNode nodeConfigurationOtherSetting = doc.SelectSingleNode("/RouterATE/Configuration/OtherSetting");
            //try
            //{
            //    string RemoteServerIp = nodeConfigurationOtherSetting.SelectSingleNode("RemoteServerIp").InnerText;                
            //    string RemoteServerPort = nodeConfigurationOtherSetting.SelectSingleNode("RemoteServerPort").InnerText;
            //    string ClientCardIp = nodeConfigurationOtherSetting.SelectSingleNode("ClientCardIp").InnerText;
            //    string ChariotPath = nodeConfigurationOtherSetting.SelectSingleNode("ChariotPath").InnerText;
            //    string TestConditionFileName = nodeConfigurationOtherSetting.SelectSingleNode("TestConditionFileName").InnerText;


            //    Debug.WriteLine("RemoteServerIp: " + RemoteServerIp);
            //    Debug.WriteLine("RemoteServerPort: " + RemoteServerPort);
            //    Debug.WriteLine("ClientCardIp: " + ClientCardIp);
            //    Debug.WriteLine("ChariotPath: " + ChariotPath);
            //    Debug.WriteLine("TestConditionFileName: " + TestConditionFileName);
                
            //    mtbRouterIntegrationConfigurationRemoteControlServerIpAddress.Text = RemoteServerIp;
            //    nudRouterIntegrationConfigurationRemoteControlServerPort.Text = RemoteServerPort;
            //    mtbRouterIntegrationConfigurationClientCardIpAddress.Text = ClientCardIp;
            //    txtRouterIntegrationConfigurationChariotFolder.Text = ChariotPath;
            //    txtRouterIntegrationConfigurationTestConditionExcelFileName.Text = TestConditionFileName;
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine("/RouterATE/Configuration/OtherSetting " + ex);
            //}

            //// Radius Server Section
            //XmlNode nodeConfigurationRadiusServerSetting = doc.SelectSingleNode("/RouterATE/Configuration/RadiusServerSetting");
            //try
            //{
            //    string RadiusServerIp = nodeConfigurationRadiusServerSetting.SelectSingleNode("RadiusServerIp").InnerText;
            //    string RadiusServerPort = nodeConfigurationRadiusServerSetting.SelectSingleNode("RadiusServerPort").InnerText;
                

            //    Debug.WriteLine("RadiusServerIp: " + RadiusServerIp);
            //    Debug.WriteLine("RadiusServerPort: " + RadiusServerPort);                
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine("/RouterATE/Configuration/RadiusServerSetting " + ex);
            //}     


            // End of Read Function Test configuration settings

            return true;
        }

        #endregion

        /*============================================================================================*/
        /*====================================== For Test Area ======================================*/
        /*============================================================================================*/
        #region

                private void ForTestRouterIntegrationTest()
                {
                    string str = string.Empty;



        //            //bRouterIntegrationTestFTThreadRunning = true;
        //            //string sFile = "";
        //            //RouterIntegrationTestMainFunction_Sanity(sFile);

        //            for (int i = 0; i < RouterIntegrationDutDataSets.Length; i++)
        //            {
        //                RouterIntegrationDutDataSets[i].TestFinish = false;

        //                //RouterIntegrationDutDataSets[i].snmp.GetMibNextTest(".1.2", ref str );
        //                //string sFileName = @"E:\MibWalk_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm") + ".txt";

        //                //RouterIntegrationDutDataSets[i].snmp.GetMibWalkPrivate(".1.3.6.1.4.1.46366.4292.77", ref str, sFileName);
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


        //            //if (chkRouterIntegrationConfigurationUseSwitch.Checked)
        //            //{
        //            //    ///* Create COM port 2 for switch control */
        //            //    comPortRouterIntegrationSwtich = new Comport2();

        //            //    if (comPortRouterIntegrationSwtich.isOpen() != true)
        //            //    {
        //            //        MessageBox.Show("COM port 2 is not ready!");
        //            //        //System.Windows.Forms.Cursor.Current = Cursors.Default;
        //            //        //btnRouterIntegrationFunctionTestRun.Enabled = true;
        //            //        return;
        //            //    }






        //            //    int x = 0;
        //            //}


        //            /* Create sub folder for saving the report data */
        //            //m_RouterIntegrationTestSubFolder = createRouterIntegrationTestSubFolder(txtRouterIntegrationFunctionTestName.Text);

        //            //string savePath = createRouterIntegrationTestSavePath(m_RouterIntegrationTestSubFolder,  m_LoopRouterIntegrationTest);
        //            //string savePath = string.Empty;

        //            //initialExcelRouterIntegrationTest(savePath);

        //            /* Save end time */
        //            //excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
        //            //saveExcelRouterIntegrationTest(savePath);


        //            /* Close excel */
        //            //closeExcelRouterIntegrationTest();
        //        }
        //        private void TestSwitch()
        //        {
        //            string str = string.Empty;

        //            if (chkRouterIntegrationConfigurationUseSwitch.Checked)
        //            {
        //                ///* Create COM port 2 for switch control */
        //                comPortRouterIntegrationSwtich = new Comport2();

        //                if (comPortRouterIntegrationSwtich.isOpen() != true)
        //                {
        //                    MessageBox.Show("COM port 2 is not ready!");
        //                    //System.Windows.Forms.Cursor.Current = Cursors.Default;
        //                    //btnRouterIntegrationFunctionTestRun.Enabled = true;
        //                    return;
        //                }

        //                switch_sf302 sf302 = new switch_sf302(comPortRouterIntegrationSwtich);

        //                str = "Login Switch sf302";
        //                SetText(str, txtRouterIntegrationFunctionTestInformation);
        //                sf302.login("cisco", "sqa11063", ref str);
        //                SetText(str, txtRouterIntegrationFunctionTestInformation);

        //                str = "Turn On Port 2";
        //                SetText(str, txtRouterIntegrationFunctionTestInformation);
        //                sf302.SwitchEthernetPort(2, true, ref str);

        //                str = "Loginout Switch sf302";
        //                SetText(str, txtRouterIntegrationFunctionTestInformation);
        //                sf302.logout(ref str);
        //                SetText(str, txtRouterIntegrationFunctionTestInformation);
        //            }
                }

        #endregion

        /*============================================================================================*/
        /*========================================  The End ==========================================*/
        /*============================================================================================*/























//        //RouterIntegrationDutDataSets[] RouterIntegrationAllDutDataSets;
//        RouterIntegrationDutDataSets[] RouterIntegrationDutDataSets;
//        RouterIntegrationVeriwave RouterIntegrationVeriwave;
        
//        /* Declare Veriwave related parameter */
//        bool bRouterIntegrationThroughputTest = false;  // Record if the test need to test throughput
//        Thread thread_RouterIntegrationVeriwaveFT;
//        //bool bRouterIntegrationTestFTThreadRunning = false;

//        /* Data Sets Function Test */
//        

//        /* */
//        //Thread theadRouterIntegrationTestStart;

//        //Excel.Workbook excelWorkBookRouterIntegrationVeriwave;
//        //Excel.Worksheet excelWorkSheetRouterIntegrationVeriwave;
//        //Excel.Range excelRangeRouterIntegrationVeriwave;
        
//        /* Time parameter to check the RD ftp server */
//        int i_SetCheckHourTime = 5;
//        int i_SetCheckMinuteTime = 00;
//        int i_SetCheckSecondTime = 00;
//        System.Timers.Timer t_Timer;

//        int i_RouterIntegrationConditionParamterNum = 30;
//        int i_RouterIntegrationCurrentIndex = 0;
//        int i_RouterIntegrationCurrentDateSetsIndex = 0;
//        int i_RouterIntegrationTestLoop = 0;
//        //Thread threadRouterIntegrationTestRunCondition;
//        string[] sa_ReportRouterIntegration;
//        string s_RouterIntegrationFinalReportInfo;

//        // 寄出報告
//        string s_EmailSenderID = string.Empty;
//        string s_EmailSenderPassword = string.Empty;
//        string[] s_EmailReceivers = null;
        
//        //CbtTlvBlderControl c_Tlvblder = null;

//        /* Excel object */
//        //Excel.Application excelAppRouterIntegrationTest;
//        //Excel.Workbook[] excelWorkBookRouterIntegrationTest;
//        //Excel.Worksheet[] excelWorkSheetRouterIntegrationTest;
//        //Excel.Range[] excelRangeRouterIntegrationTest;


        
        
//        string[,] sa_TestConditionSets;



        

//        //string[,] saa_testConfigSanityRouterIntegrationTest;
//        //string[,] saa_testConfigFullRouterIntegrationTest;
//        string[,] testConfigRouterIntegrationTest;
//        string[,] testConfigRouterIntegrationSanityTest;
        
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
//        int i_RouterIntegrationTestIndexMaximun = 1000;
//        //List<CbtMibAccess> ls_CurrentMibs = new List<CbtMibAccess>();

//        //string s_ErrorMsgForReport;          

//        //bool b_CfgNotExist = false;
             
              




//        private void DoRouterIntegrationTestStart()
//        {
//            string str = string.Empty;

//            //string RdPath = "192.168.65.201";
//            string RdPath = txtRouterIntegrationConfigurationRdFtpServerPath.Text;
//            string RdTempPath = System.Windows.Forms.Application.StartupPath + "\\Temp";
//            //string sDirOrig = "/Technicolor/Taipei_BFC5.7.1mp3_RG/CGA2121";
//            string sDirOrig = "/Technicolor/";
//            if (rbtnRouterIntegrationConfigurationRdFtpMp3Folder.Checked)
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

//            m_RouterIntegrationTestSubFolder = createRouterIntegrationTestSubFolder(txtRouterIntegrationFunctionTestName.Text);
//            //string saveLogPath = m_RouterIntegrationTestSubFolder + "\\" + "PreProcess.txt";
//            string saveLogPath ;

//            for (int i = 0; i < RouterIntegrationDutDataSets.Length; i++)
//            {
//                i_RouterIntegrationCurrentDateSetsIndex = i;

//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    bRouterIntegrationTestFTThreadRunning = false;
//                    //MessageBox.Show("Abort test", "Error");
//                    //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                    //threadRouterIntegrationTestFT.Abort();
//                    return;
//                    // Never go here ...
//                }

//                if (RouterIntegrationDutDataSets[i].TestFinish)
//                {
//                    continue;
//                }

//                /* Create folder for each test or sku */
//                if (RouterIntegrationDutDataSets[i].TestType.ToLower().IndexOf("throughput") >= 0)
//                { //Reset Device by SNMP                     

//                    bool isExists;
//                    string subPath = m_RouterIntegrationTestSubFolder + "\\Veriwave";                   
//                    RouterIntegrationDutDataSets[i].ReportSavePath = subPath;

//                    isExists = System.IO.Directory.Exists(subPath);
//                    if (!isExists)
//                    {
//                        str = "Create Veriwave Folder.";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        System.IO.Directory.CreateDirectory(subPath);
//                        str = "Done.";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }                    
//                }
//                else
//                {
//                    /*  Create Sub-folder in terms of sku model */                    

//                    bool isExists;
//                    string subPath = m_RouterIntegrationTestSubFolder + "\\" + RouterIntegrationDutDataSets[i].SkuModel;
//                    RouterIntegrationDutDataSets[i].ReportSavePath = subPath;

//                    isExists = System.IO.Directory.Exists(subPath);
//                    if (!isExists)
//                    {
//                        str = "Create SKU Model Folder: " + RouterIntegrationDutDataSets[i].SkuModel;
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        System.IO.Directory.CreateDirectory(subPath);
//                        str = "Done.";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }                    
//                }

//                //// for temp test
//                //string ExcelFile = txtRouterIntegrationConfigurationTestConditionExcelFile.Text;
//                //RouterIntegrationTestMainFunction(ExcelFile);
//                //RouterIntegrationDutDataSets[i].TestFinish = true;                                

//                saveLogPath = RouterIntegrationDutDataSets[i].ReportSavePath + "\\" + "HeaderLog_" + DateTime.Now.ToString("yyyyMMdd-hhmmss")+ ".txt";
                
//                str = "Current SKU is: " + RouterIntegrationDutDataSets[i].SkuModel;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = "Check if folder exists in RD server today.";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                /* Check if FW is in RD server */
//                sDir = sDirOrig;
//                //sDir = sDir + "/CGA2121_" + sSku[i] + "/";
//                if (RouterIntegrationDutDataSets[i].SkuModel.IndexOf("GERMANY") >= 0)
//                {
//                    sDir = sDir + "/CGA2121_GERMANY/";
//                }
//                else
//                {
//                    sDir = sDir + "/CGA2121_" + RouterIntegrationDutDataSets[i].SkuModel + "/";
//                }
//                sDir = RdPath + sDir;

//                str = "Ftp Path is: ftp://" + sDir;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                        File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                        str = String.Format("Sku: {0} 找不到今天的資料夾", RouterIntegrationDutDataSets[i].SkuModel);

//                        //寄出報告
//                        try
//                        {
//                            SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                        }
//                        catch (Exception ex)
//                        {
//                            str = "Send Mail Failed: " + ex.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        }

//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
//                        continue;
//                    }

//                    str = "The Date Folder Exists: ftp://" + sDir;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    string sDownloadFileName = RouterIntegrationDutDataSets[i].FwFileName;
//                    sDownloadFileName = sDownloadFileName.Replace("#DATE", sShorDay);
//                    RouterIntegrationDutDataSets[i].CurrentFwName = sDownloadFileName;

//                    str = "Check if file exists in RD server today:" + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    ftpclient = null;
//                    ftpclient = new CbtFtpClient(sDir, "cbtsqa", "jenkins");

//                    /* Get File list in today's folder */
//                    list.Clear();
//                    list = ftpclient.GetFtpFileList();
//                    string[] sFileName = new string[20];
//                    int iIndex = 0;
                    
//                    //列出目錄中所有檔案
//                    str = "List all the files in Folder: "  + sDir;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    foreach (string s in list)
//                    {
//                        string[] sFiles = s.Split(' ');
//                        foreach (string st in sFiles)
//                        {
//                            if (st.IndexOf("CGA") >= 0)
//                            {
//                                Invoke(new SetTextCallBackT(SetText), new object[] { st, txtRouterIntegrationFunctionTestInformation });
//                                sFileName[iIndex++] = st;
//                            }
//                        }
//                    }
                    
//                    //確認檔案是否存在
//                    str = "Check if the file exists :" + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                        File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                        str = String.Format("Sku: {0} 找不到FW檔案", RouterIntegrationDutDataSets[i].SkuModel);

//                        //寄出報告
//                        try
//                        {
//                            SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                        }
//                        catch (Exception ex)
//                        {
//                            str = "Send Mail Failed: " + ex.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        }

//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
//                        continue;
//                    }

//                    str = "Start to Download Fw from RD server: " + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    ftpclient = null;
//                    ftpclient = new CbtFtpClient(RdPath, "cbtsqa", "jenkins");
//                    str = "Start to Download File: " + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = "";
//                    ftpclient.DownloadFile(RdTempPath, sDir, sDownloadFileName, ref str);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = "Download File Succeed: " + sDownloadFileName;
//                    sInTempFile[i] = sDownloadFileName;

//                    ftpclient = null;
//                }
//                catch (Exception ex)
//                {
//                    str = "RD FW Doanload Exception: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    saveLogPath = m_RouterIntegrationTestSubFolder + "\\" + "PreProcess_" + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".txt";

//                    File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Sku: {0} Download FW from RD server Failed", RouterIntegrationDutDataSets[i].SkuModel);

//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                    }
//                    catch (Exception exp)
//                    {
//                        str = "Send Mail Failed: " + exp.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }

//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
//                    continue;
//                }

//                // RD FW 下載完成  
//                str = "RD FW Doanload Finished!!! Start to Upload to NFS Server!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                string NfsPath = txtRouterIntegrationConfigurationNfsServerPath.Text;
//                //上傳 Firmware 檔案到NFS Server 上面去.

//                str = "Check if Local tftp server exists?";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                if (!Directory.Exists(NfsPath))
//                { //連不到NFS的資料夾,等一個小時之後再試
//                    str = "Local Tftp Server is unreachable!!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    saveLogPath = m_RouterIntegrationTestSubFolder + "\\" + "PreProcess_" + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".txt";

//                    File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("NFS server unreachable");

//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                    }
//                    catch (Exception exp)
//                    {
//                        str = "Send Mail Failed: " + exp.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }

//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
//                    continue;
//                }

//                str = "Server exist. Copy file to local tftp server: ";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                string sFwFile = RdTempPath + "\\" + RouterIntegrationDutDataSets[i].CurrentFwName;

//                str = "Copy File to Local NFS server: " + sFwFile;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                string sDestFile = NfsPath + "\\" + RouterIntegrationDutDataSets[i].CurrentFwName;

//                //CopyFileToNfsFoler(sFwFile, NfsPath, true);
//                File.Copy(sFwFile, sDestFile, true);

//                str = "Copy File Succeed.";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //刪除已下載FW
//                //File.Delete(sFwFile);

//                str = "Finished!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = "Start to Write FW File Name to each Device Process!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //將要Download的 FW 檔名寫入Device中
//                int snmpVersion = rbtnRouterIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnRouterIntegrationConfigurationSnmpV2.Checked ? 2 : 3;

//                string sIp = RouterIntegrationDutDataSets[i].IpAddress;

//                str = "Ping Device First: " + sIp;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                 //Ping Device First
//                /* Ping ip first */
//                //if (!QuickPingRouterIntegrationTest(mtbRouterIntegrationFunctionTestCmIpAddress.Text, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//                if (!QuickPingRouterIntegrationTest(sIp, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//                {
//                    str = " Failed!! ";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    saveLogPath = m_RouterIntegrationTestSubFolder + "\\" + "PreProcess_" + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".txt";

//                    File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Sku: {0} Ping Device Failed.", RouterIntegrationDutDataSets[i].SkuModel);

//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                    }
//                    catch (Exception exp)
//                    {
//                        str = "Send Mail Failed: " + exp.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }

//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
//                    continue;
//                }

//                //Setting Tftp Server IP
//                str = "Setting Tftp Server IP to Device: " + RouterIntegrationDutDataSets[i].SkuModel + ", " + RouterIntegrationDutDataSets[i].CurrentFwName;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                str = "Tftp Server IP: " + mtbRouterIntegrationConfigurationServerIp.Text;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerOid.Text.Trim(), mibType, mtbRouterIntegrationConfigurationServerIp.Text.Trim(), snmpVersion);
//                Thread.Sleep(1000);
//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //Setting Tftp Server Address
//                str = "Setting Tftp Server Address Type: " + txtRouterIntegrationConfigurationTftpServerTypeValue.Text;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerTypeOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerTypeOID.Text.Trim(), mibType, txtRouterIntegrationConfigurationTftpServerTypeValue.Text.Trim(), snmpVersion);
//                Thread.Sleep(1000);
//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerTypeOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = "Write Fw File Name to Device: " + RouterIntegrationDutDataSets[i].SkuModel + ", " + RouterIntegrationDutDataSets[i].CurrentFwName;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationImageFileOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationImageFileOid.Text.Trim(), mibType, RouterIntegrationDutDataSets[i].CurrentFwName, snmpVersion);
//                Thread.Sleep(1000);

//                // 讀回FW Name
//                valueResponsed = "";
//                mibType = "";

//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationImageFileOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(sFwFileNameOid, "OctetString", sFwFiles[i], snmpVersion);
//                //setStatus = Snmp_CableMibReadWriteTest.SetMibSingleValueByVersion(sOid, mibType, sWriteValue, snmpVersion);              

//                str = "Download FW and Reboot By SNMP";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                // Set Admin Status to 1 for downloading firmware
//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), mibType, txtRouterIntegrationConfigurationAdminStatusValue.Text.Trim(), snmpVersion);
//                Thread.Sleep(1000);
//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                
//                if (RouterIntegrationDutDataSets[i].TestType.ToLower().IndexOf("throughput") >= 0)
//                { //Reset Device by SNMP 
//                    Thread.Sleep(2000);

//                    str = "Before Reset, Check Device Update (Inprocess) Status: ";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    valueResponsed = "";
//                    mibType = "";
//                    RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = "Wait " + nudRouterIntegrationConfigurationFwUpdateWaitTime.Value.ToString() + " Seconds for FW Update";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    stopWatch.Stop();
//                    stopWatch.Reset();
//                    stopWatch.Restart();

//                    /* Wait for Fw update finish */
//                    while (true)
//                    {
//                        if (bRouterIntegrationTestFTThreadRunning == false)
//                        {
//                            return;
//                            // Never go here ...
//                        }

//                        if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationConfigurationFwUpdateWaitTime.Value * 1000))
//                            break;

//                        //Thread.Sleep(1000);
//                        Thread.Sleep(1000);
//                        str = String.Format(".");
//                        Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }

//                    str = "Check In-Process Status finished:";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    stopWatch.Stop();
//                    stopWatch.Reset();
//                    stopWatch.Restart();

//                    while (true)
//                    {
//                        bool inProcessStatus = true;

//                        if (bRouterIntegrationTestFTThreadRunning == false)
//                        {
//                            return;
//                            // Never go here ...
//                        }

//                        if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationConfigurationInProcessTimeout.Value * 1000))
//                            break;

//                        str = ".";
//                        Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                        try
//                        {
//                            valueResponsed = "";
//                            mibType = "";
//                            RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//                            str = "Mib Type: " + mibType;
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                            str = String.Format("Read Value: {0}", valueResponsed);
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                            if (str != "3") inProcessStatus = false;
//                        }
//                        catch (Exception ex)
//                        {
//                            str = "Get In-Process Error: " + ex.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        }

//                        if (inProcessStatus) break;
//                        Thread.Sleep(1000);
//                    }

//                    //Wait Time for Device Reset
//                    str = "Wait For Device Reset... " + nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString() + " Seconds.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    stopWatch.Stop();
//                    stopWatch.Reset();
//                    stopWatch.Restart();

//                    while (true)
//                    {
//                        if (bRouterIntegrationTestFTThreadRunning == false)
//                        {
//                            bRouterIntegrationTestFTThreadRunning = false;
//                            return;
//                            // Never go here ...
//                        }

//                        if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationFunctionTestDutRebootTimeout.Value * 1000))
//                        {
//                            break;
//                        }

//                        if(stopWatch.ElapsedMilliseconds % (1000*60) == 0)
//                        {
//                            //str = String.Format("{0} Seconds Passed.", stopWatch.ElapsedMilliseconds % (1000));
//                            //Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        }
//                        else
//                        { 
//                            Thread.Sleep(1000);
//                            str = String.Format(".");
//                            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        }
//                    }

//                    str = "Done." + nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString();

//                    // 讀回目前的 Description 以確認 FW 版本正確
//                    str = String.Format("Get Description to check if FW correct.");
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    valueResponsed = "";
//                    mibType = "";

//                    RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationDescriptionOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    //確認FW版本中有要測試的日期

//                    string sFwFileName = RouterIntegrationDutDataSets[i].CurrentFwName;
//                    sFwFileName = sFwFileName.Substring(0, sFwFileName.Length - 4);

//                    str = "FW File Name should be: " + sFwFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });     

//                    if (valueResponsed.IndexOf(sFwFileName) < 0)
//                    {
//                        str = " Failed!! ";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                        File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                        str = String.Format("Sku: {0} FW File Name of Dut is wrong.", RouterIntegrationDutDataSets[i].SkuModel);

//                        //寄出報告
//                        try
//                        {
//                            SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                        }
//                        catch (Exception exp)
//                        {
//                            str = "Send Mail Failed: " + exp.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        }

//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
//                        continue;                       
//                    }

//                    str = " succeed!! ";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = "Save Log to HeaderLog.txt";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = "Done." + nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString();

//                    File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                    RouterIntegrationVeriwaveMainFunction(i);
//                    RouterIntegrationDutDataSets[i].TestFinish = true;
//                }
//                else
//                {//Read by Comport
//                    /* Set Comport to device setting comport */
//                    str = "Set Comport to : " + RouterIntegrationDutDataSets[i].ComportNum;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    comPortRouterIntegrationTest.Close();
//                    comPortRouterIntegrationTest.SetPortName(RouterIntegrationDutDataSets[i].ComportNum);
//                    comPortRouterIntegrationTest.Open();

//                    if (comPortRouterIntegrationTest.isOpen() != true)
//                    {
//                        str = " Failed!! ";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });                       

//                        File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                        str = String.Format("Sku: {0} Comport: {1} Open Failed.", RouterIntegrationDutDataSets[i].SkuModel, RouterIntegrationDutDataSets[i].ComportNum);

//                        //寄出報告
//                        try
//                        {
//                            SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, str, str, saveLogPath);
//                        }
//                        catch (Exception exp)
//                        {
//                            str = "Send Mail Failed: " + exp.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        }

//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
//                        continue;                       
//                    }

//                    str = "Done.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    stopWatch.Stop();
//                    stopWatch.Reset();
//                    stopWatch.Restart();

//                    /* Wait for Fw update finish */
//                    while (true)
//                    {
//                        if (bRouterIntegrationTestFTThreadRunning == false)
//                        {
//                            return;
//                            // Never go here ...
//                        }

//                        if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationConfigurationFwUpdateWaitTime.Value * 1000))
//                            break;

//                        //if (stopWatch.ElapsedMilliseconds % (1000 * 60) == 0)
//                        //{
//                        //    str = String.Format("====ATE MSG ==== :{0} Seconds Passed.", stopWatch.ElapsedMilliseconds % (1000));
//                        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        //}

//                        //Thread.Sleep(1000);
//                        str = comPortRouterIntegrationTest.ReadLine();

//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }

//                    str = "Check Device Update (Inprocess) Status After FW update: ";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    valueResponsed = "";
//                    mibType = "";
//                    RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    //Wait Time for Device Reset
//                    str = "Wait For Device Reset... " + nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString() + " Seconds.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    stopWatch.Stop();
//                    stopWatch.Reset();
//                    stopWatch.Restart();

//                    while (true)
//                    {
//                        if (bRouterIntegrationTestFTThreadRunning == false)
//                        {
//                            bRouterIntegrationTestFTThreadRunning = false;
//                            return;
//                            // Never go here ...
//                        }

//                        if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationFunctionTestDutRebootTimeout.Value * 1000))
//                        {
//                            break;
//                        }

//                        //if (stopWatch.ElapsedMilliseconds % (1000 * 60) == 0)
//                        //{
//                        //    str = String.Format("====ATE MSG ==== :{0} Seconds Passed.", stopWatch.ElapsedMilliseconds % (1000));
//                        //    Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        //}

//                        str = comPortRouterIntegrationTest.ReadLine();

//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }

//                    str = "Done." + nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString();

//                    if (!CGA2121CheckFwFileName_RouterIntegrationTest(i))
//                    { //Read show cfg failed, or the FW filename is wrong 

//                        str = "Show version failed, or the FW filename is wrong";
//                        Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        File.AppendAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);

//                        //寄出報告
//                        try
//                        {
//                            SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Download FW Process Failed", "Copy Rd FW to NFS server Process Failed!!!", saveLogPath);
//                        }
//                        catch (Exception ex)
//                        {
//                            str = "Send Mail Failed: " + ex.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        }

//                        continue;
//                    }

//                    str = "Save Log to HeaderLog.txt";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
                    
//                    i_RouterIntegrationCurrentDateSetsIndex = i;
//                    string ExcelFile = txtRouterIntegrationConfigurationTestConditionExcelFile.Text;

//                    //str = "Read Excel Test Condition!!";
//                    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    //if (!File.Exists(ExcelFile))
//                    //{
//                    //    str = "Excel File doesn't Exist: " + ExcelFile;
//                    //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    //    str = "Read next datagridview test condition.";
//                    //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                        
//                    //    continue;
//                    //}

//                    //RouterIntegrationTestMainFunction(ExcelFile);
//                    RouterIntegrationTestMainFunction_Sanity(ExcelFile);
//                    RouterIntegrationDutDataSets[i].TestFinish = true;
//                }                
//            }

//            i_RouterIntegrationTestLoop++;                
            
//            bool bAllTestFinished = true;
//            for (int i = 0; i < RouterIntegrationDutDataSets.Length; i++)
//            {// Check if all the device test finished?
//                if (!RouterIntegrationDutDataSets[i].TestFinish)
//                    bAllTestFinished = false;
//            }

//            if (bAllTestFinished)
//            {
//                // 重新設定下載FW時間設定為明天早上5:00
//                i_SetCheckHourTime = 5;
//                i_SetCheckMinuteTime = 00;
//                i_SetCheckSecondTime = 00;

//                str = "All Device Test Finished!!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });                

//                str = " ===== Test Completed!! =====";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = "Wait another day to Test and Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = "Waiting for time up....";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                return;
//            }

//            if (DateTime.Now.Hour > 18 || i_RouterIntegrationTestLoop >= 6)
//            {
//                // 重新設定下載FW時間設定為明天早上5:00
//                i_SetCheckHourTime = 5;
//                i_SetCheckMinuteTime = 00;
//                i_SetCheckSecondTime = 00;

//                str = "Some Device Test Failed!!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //str = " ===== Test Completed!! =====";
//                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = "Wait another day to Test and Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = "Waiting for time up....";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            }

//            if (!bAllTestFinished)
//            { //Some Device Test Failed.
//                //連不到NFS的資料夾,等一個小時之後再試
//                str = "Some Device Test Failed, Wait for one hour and try again!!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                return;
//            }
            

//            //// 準備開始Device測試, 下次下載FW時間設定為明天早上5:00
//            //i_SetCheckHourTime = 5;
//            //i_SetCheckMinuteTime = 00;
//            //i_SetCheckSecondTime = 00;



//            //// Check RD ftp Server if the FW driver exist
//            //str = "Download FW from RD server to NFS Server.";
//            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            ////if (!CopyRdFwToNfsServerRouterIntegrationTest())
//            ////{
//            ////    str = "Copy Rd FW to NFS server Process Failed!!! Wait next Test Time: " + i_SetCheckHourTime + ":" + i_SetCheckMinuteTime + ":" + i_SetCheckSecondTime; ;
//            ////    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            ////    File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//            ////    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//            ////    //寄出報告
//            ////    try
//            ////    {
//            ////        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Download FW Process Failed", "Copy Rd FW to NFS server Process Failed!!!", saveLogPath);
//            ////    }
//            ////    catch (Exception ex)
//            ////    {
//            ////        str = "Send Mail Failed: " + ex.ToString();
//            ////        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            ////    }

//            ////    str = "Setting Time is: " + i_SetCheckHourTime + ":" + i_SetCheckMinuteTime + ":" + i_SetCheckSecondTime;
//            ////    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            ////    str = "Wait Time ....";
//            ////    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            ////    return;
//            ////    //theadRouterIntegrationTestStart.Abort();
//            ////}

//            ////str = "Download FW Succeed!!";
//            ////Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            ////File.WriteAllText(saveLogPath, txtRouterIntegrationFunctionTestInformation.Text);
//            ////Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//            ///* Reset Device */
//            //for (int i = 0; i < RouterIntegrationDutDataSets.Length; i++)
//            //{
//            //    //string valueResponsed = string.Empty;
//            //    //string expectedValue = string.Empty;
//            //    //string mibType = string.Empty;
//            //    string saveLogSubPath = m_RouterIntegrationTestSubFolder + "\\" + RouterIntegrationDutDataSets[i].SkuModel + "\\" + "HeaderLog.txt";
//            //    string temp = RouterIntegrationDutDataSets[i].ReportSavePath + "\\" + "HeaderLog.txt";
                
//            //    if (bRouterIntegrationTestFTThreadRunning == false)
//            //    {
//            //        bRouterIntegrationTestFTThreadRunning = false;
//            //        //MessageBox.Show("Abort test", "Error");
//            //        //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//            //        //threadRouterIntegrationTestFT.Abort();
//            //        return;
//            //        // Never go here ...
//            //    }

//            //    if (RouterIntegrationDutDataSets[i].TestType.ToLower().IndexOf("throughput") >= 0)
//            //    { //Reset Device by SNMP 
//            //        str = "Create Veriwave Folder.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        bool isExists;
//            //        string subPath = m_RouterIntegrationTestSubFolder + "\\Veriwave";
//            //        saveLogSubPath = subPath + "\\" + "HeaderLog.txt";
//            //        RouterIntegrationDutDataSets[i].ReportSavePath = subPath;

//            //        isExists = System.IO.Directory.Exists(subPath);
//            //        if (!isExists)
//            //            System.IO.Directory.CreateDirectory(subPath);

//            //        str = "Done.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        str = "Download FW and Reboot By SNMP";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        // Download FW and Reboot By SNMP

//            //        int snmpVersion = rbtnRouterIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnRouterIntegrationConfigurationSnmpV2.Checked ? 2 : 3;
//            //        //
//            //        // Set Admin Status to 1 for downloading firmware
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//            //        RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), mibType, txtRouterIntegrationConfigurationAdminStatusValue.Text.Trim(), snmpVersion);
//            //        Thread.Sleep(1000);
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//            //        str = "Mib Type: " + mibType;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        //RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationResetOid.Text.Trim(), dudRouterIntegrationConfigurationResetType.Text, txtRouterIntegrationConfigurationResetValue.Text, snmpVersion);
//            //        Thread.Sleep(2000);

//            //        str = "Before Reset, Check Device Update (Inprocess) Status: ";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //        str = "Mib Type: " + mibType;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        str = "Wait " + nudRouterIntegrationConfigurationFwUpdateWaitTime.Value.ToString() + " Seconds for FW Update";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        stopWatch.Stop();
//            //        stopWatch.Reset();
//            //        stopWatch.Restart();

//            //        /* Wait for Fw update finish */
//            //        while (true)
//            //        {
//            //            if (bRouterIntegrationTestFTThreadRunning == false)
//            //            {
//            //                return;
//            //                // Never go here ...
//            //            }

//            //            if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationConfigurationFwUpdateWaitTime.Value * 1000))
//            //                break;

//            //            //Thread.Sleep(1000);
//            //            Thread.Sleep(1000);
//            //            str = String.Format(".");
//            //            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //        }

//            //        str = "Check In-Process Status finished:";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        stopWatch.Stop();
//            //        stopWatch.Reset();
//            //        stopWatch.Restart();

//            //        while (true)
//            //        {
//            //            bool inProcessStatus = true;

//            //            if (bRouterIntegrationTestFTThreadRunning == false)
//            //            {
//            //                return;
//            //                // Never go here ...
//            //            }

//            //            if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationConfigurationInProcessTimeout.Value * 1000))
//            //                break;

//            //            str = ".";
//            //            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //            try
//            //            {
//            //                valueResponsed = "";
//            //                mibType = "";
//            //                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //                str = "Mib Type: " + mibType;
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //                str = String.Format("Read Value: {0}", valueResponsed);
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //                if (str != "3") inProcessStatus = false;
//            //            }
//            //            catch (Exception ex)
//            //            {
//            //                str = "Get In-Process Error: " + ex.ToString();
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //            }

//            //            if (inProcessStatus) break;
//            //            Thread.Sleep(1000);
//            //        }

//            //        //Wait Time for Device Reset
//            //        str = "Wait For Device Reset... " + nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString() + " Seconds.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        stopWatch.Stop();
//            //        stopWatch.Reset();
//            //        stopWatch.Restart();

//            //        while (true)
//            //        {
//            //            if (bRouterIntegrationTestFTThreadRunning == false)
//            //            {
//            //                bRouterIntegrationTestFTThreadRunning = false;
//            //                return;
//            //                // Never go here ...
//            //            }

//            //            if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationFunctionTestDutRebootTimeout.Value * 1000))
//            //            {
//            //                break;
//            //            }

//            //            Thread.Sleep(1000);
//            //            str = String.Format(".");
//            //            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //        }

//            //        str = "Done." + nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString();

//            //        str = "Save Log to HeaderLog.txt";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                                       
//            //        str = "Done." + nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString();

//            //        File.WriteAllText(saveLogSubPath, txtRouterIntegrationFunctionTestInformation.Text);
//            //        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
//            //    }
//            //    else
//            //    { //Reset by Comport

//            //        /*  Create Sub-folder in terms of sku model */
//            //        str = "Create SKU Model Folder: " + RouterIntegrationDutDataSets[i].SkuModel;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        bool isExists;
//            //        string subPath = m_RouterIntegrationTestSubFolder + "\\" + RouterIntegrationDutDataSets[i].SkuModel;
//            //        RouterIntegrationDutDataSets[i].ReportSavePath = subPath;

//            //        isExists = System.IO.Directory.Exists(subPath);
//            //        if (!isExists)
//            //            System.IO.Directory.CreateDirectory(subPath);

//            //        str = "Done.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        /* Set Comport to device setting comport */
//            //        str = "Set Comport to : " + RouterIntegrationDutDataSets[i].ComportNum;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        comPortRouterIntegrationTest.Close();
//            //        comPortRouterIntegrationTest.SetPortName(RouterIntegrationDutDataSets[i].ComportNum);
//            //        comPortRouterIntegrationTest.Open();

//            //        if (comPortRouterIntegrationTest.isOpen() != true)
//            //        {
//            //            //MessageBox.Show("COM port is not ready!");
//            //            //System.Windows.Forms.Cursor.Current = Cursors.Default;
//            //            //btnRouterIntegrationFunctionTestRun.Enabled = true;
//            //            return;
//            //        }

//            //        str = "Done.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        str = "Download FW and Reboot By SNMP";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        // Download FW and Reboot By SNMP
//            //        //int snmpVersion = rbtnRouterIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnRouterIntegrationConfigurationSnmpV2.Checked ? 2 : 3;
//            //        //RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationResetOid.Text.Trim(), dudRouterIntegrationConfigurationResetType.Text, txtRouterIntegrationConfigurationResetValue.Text, snmpVersion);

//            //        int snmpVersion = rbtnRouterIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnRouterIntegrationConfigurationSnmpV2.Checked ? 2 : 3;
//            //        //
//            //        // Set Admin Status to 1 for downloading firmware
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//            //        RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), mibType, txtRouterIntegrationConfigurationAdminStatusValue.Text.Trim(), snmpVersion);
//            //        Thread.Sleep(1000);
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//            //        str = "Mib Type: " + mibType;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        //RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationResetOid.Text.Trim(), dudRouterIntegrationConfigurationResetType.Text, txtRouterIntegrationConfigurationResetValue.Text, snmpVersion);
//            //        Thread.Sleep(2000);

//            //        str = "Before Reset, Check Device Update (Inprocess) Status: ";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //        str = "Mib Type: " + mibType;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                    
//            //        str = "Wait " + nudRouterIntegrationConfigurationFwUpdateWaitTime.Value.ToString() + " Seconds for FW Update";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        stopWatch.Stop();
//            //        stopWatch.Reset();
//            //        stopWatch.Restart();

//            //        /* Wait for Fw update finish */
//            //        while (true)
//            //        {
//            //            if (bRouterIntegrationTestFTThreadRunning == false)
//            //            {
//            //                return;
//            //                // Never go here ...
//            //            }

//            //            if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationConfigurationFwUpdateWaitTime.Value * 1000))
//            //                break;

//            //            //Thread.Sleep(1000);
//            //            str = comPortRouterIntegrationTest.ReadLine();

//            //            if (str != "")
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //        }

//            //        str = "Check Device Update (Inprocess) Status After FW update: ";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //        valueResponsed = "";
//            //        mibType = "";
//            //        RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //        str = "Mib Type: " + mibType;
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        //try
//            //        //{
//            //        //    valueResponsed = "";
//            //        //    mibType = "";
//            //        //    RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //        //    str = "Mib Type: " + mibType;
//            //        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        //    str = String.Format("Read Value: {0}", valueResponsed);
//            //        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        //    if (str != "3") inProcessStatus = false;
//            //        //}
//            //        //catch (Exception ex)
//            //        //{
//            //        //    str = "Get In-Process Error: " + ex.ToString();
//            //        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //        //}


//            //        //stopWatch.Stop();
//            //        //stopWatch.Reset();
//            //        //stopWatch.Restart();
                    
//            //        //while (true)
//            //        //{
//            //        //    bool inProcessStatus = true;

//            //        //    if (bRouterIntegrationTestFTThreadRunning == false)
//            //        //    {
//            //        //        return;
//            //        //        // Never go here ...
//            //        //    }

//            //        //    if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationConfigurationInProcessTimeout.Value * 1000))
//            //        //        break;

//            //        //    str = ".";
//            //        //    Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                        
//            //        //    try
//            //        //    {
//            //        //        valueResponsed = "";
//            //        //        mibType = "";
//            //        //        RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121FwUpdateStatysOid, ref valueResponsed, ref mibType, snmpVersion);
//            //        //        str = "Mib Type: " + mibType;
//            //        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        //        str = String.Format("Read Value: {0}", valueResponsed);
//            //        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        //        if (str != "3") inProcessStatus = false;                           
//            //        //    }
//            //        //    catch (Exception ex)
//            //        //    {
//            //        //        str = "Get In-Process Error: " + ex.ToString();
//            //        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //        //    }

//            //        //    if (inProcessStatus) break;
//            //        //    Thread.Sleep(1000);
//            //        //}
                    
//            //        //Wait Time for Device Reset
//            //        str = "Wait For Device Reset... " + nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString() + " Seconds.";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        stopWatch.Stop();
//            //        stopWatch.Reset();
//            //        stopWatch.Restart();

//            //        while (true)
//            //        {
//            //            if (bRouterIntegrationTestFTThreadRunning == false)
//            //            {
//            //                bRouterIntegrationTestFTThreadRunning = false;
//            //                return;
//            //                // Never go here ...
//            //            }

//            //            if (stopWatch.ElapsedMilliseconds >= Convert.ToInt32(nudRouterIntegrationFunctionTestDutRebootTimeout.Value * 1000))
//            //            {
//            //                break;
//            //            }
                        
//            //            str = comPortRouterIntegrationTest.ReadLine();

//            //            if (str != "")
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //        }

//            //        str = "Done." + nudRouterIntegrationFunctionTestDutRebootTimeout.Value.ToString();                    

//            //        str = "Save Log to HeaderLog.txt";
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //        File.WriteAllText(saveLogSubPath, txtRouterIntegrationFunctionTestInformation.Text);
//            //        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
//            //    }
//            //}
            
//            //str = "Reset Device Finished. Now Start Function Test.";
//            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //threadRouterIntegrationTestFT = new Thread(new ThreadStart(DoRouterIntegrationFunctionTest));
//            //threadRouterIntegrationTestFT.Name = "";
//            //threadRouterIntegrationTestFT.Start();        
//        }
        
//        /* Main function thread */
//        /* Report power thread function */
//        /* For reference , This function want't be used */
//        private void DoRouterIntegrationFunctionTest()
//        {
//            string str = string.Empty;            

//            int m_totalTimes = 1;  // Default test time
//            m_LoopRouterIntegrationTest = 1; // Reset loop counter
//            if (chkRouterIntegrationFunctionTestScheduleOnOff.Checked)                
//                m_totalTimes = Convert.ToInt32(nudRouterIntegrationFunctionTestTimes.Value);                       

//            /* Start Test Loop */
//            do
//            {
//                int iCount = 1 ;// indicate the test condition number running now. 
//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    bRouterIntegrationTestFTThreadRunning = false;
//                    //MessageBox.Show("Abort test", "Error");
//                    //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                    //threadRouterIntegrationTestFT.Abort();
//                    return;
//                    // Never go here ...
//                }

//                /* Read Test Condition and run each excel file as Main function */
//                for (int conditionRow = 0; conditionRow < 1; conditionRow++)
//                {
//                    if (bRouterIntegrationTestFTThreadRunning == false)
//                    {
//                        bRouterIntegrationTestFTThreadRunning = false;
//                        //MessageBox.Show("Abort test", "Error");
//                        //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                        //threadRouterIntegrationTestFT.Abort();
//                        return;
//                        // Never go here ...
//                    }

//                    string ExcelFile = txtRouterIntegrationConfigurationTestConditionExcelFile.Text;

//                    str = "Read Excel Test Condition!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    if (!File.Exists(ExcelFile))
//                    {
//                        str = "Excel File doesn't Exist: " + ExcelFile;
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        str = "Read next datagridview test condition.";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        iCount++;
//                        continue;
//                    }

//                    str = " ===== Start to run Excel item: " + iCount.ToString() + " =====";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });                   

//                    RouterIntegrationTestMainFunction(ExcelFile);               

//                    iCount++;
//                }

//                m_totalTimes--;
//                m_LoopRouterIntegrationTest++;
//            } while (m_totalTimes != 0);

//            str = " ===== Test Completed!! =====";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            str = "Wait another day to Test and Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            str = "Waiting for time up....";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            return;
//            //bRouterIntegrationTestFTThreadRunning = false;
//            //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));                
//        }

//        /* Main Function for Integration */
//        private bool RouterIntegrationTestMainFunction_Sanity(string sExcelFileName)
//        {
//            string str = string.Empty;
//            int iCount = 0; // Indicate how many excel index row has read.
//            int iIndexCount = 0; // Indicate same index has read.

//            int iIndexCurrentRun = 1;
//            string sExcelFile = string.Empty;

//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();

//            /* Initial Value */
//            m_PositionRouterIntegrationTest = 16;
//            //testConfigRouterIntegrationTest = null;

//            string subFolder = m_RouterIntegrationTestSubFolder + "\\" + RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel;
//            string savePath = createRouterIntegrationTestSaveExcelFile(subFolder, RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel, m_LoopRouterIntegrationTest, sExcelFileName);
//            string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";            
//            //string savePath = createRouterIntegrationTestSaveExcelFile(m_RouterIntegrationTestSubFolder, RouterIntegrationDutDataSets[i].SkuModel, m_LoopRouterIntegrationTest, sExcelFileName);
//            /* initial Excel component */
//            initialExcelRouterIntegrationTest(savePath);

//            try
//            {
//                /* Fill Loop, PowerLevel and Constellation in Excel */
//                xls_excelWorkSheetRouterIntegrationTest.Cells[9, 3] = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SerialNumber;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[10, 3] = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SwVersion;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[11, 3] = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].HwVersion;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[4, 9] = m_LoopRouterIntegrationTest;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[5, 9] = Path.GetFileName(sExcelFileName);
//                xls_excelWorkSheetRouterIntegrationTest.Cells[6, 9] = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[7, 9] = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].CurrentFwName;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[8, 9] = "";
                
//            }
//            catch (Exception ex)
//            {
//                Debug.WriteLine(ex);
//            }

//            ///* Read excel content of excel file */
//            //if (!ReadTestConfig_ReadExcelRouterIntegrationTest(sExcelFileName))
//            //{
//            //    str = " Failed!!";
//            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //    //string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//            //    File.WriteAllText(saveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//            //    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//            //    try
//            //    {
//            //        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = "Read File Failed";

//            //        /* Write Console Log File as HyperLink in Excel report */
//            //        xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3];
//            //        xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//            //    }
//            //    catch (Exception ex)
//            //    {
//            //        str = "Write Data to Excel File Exception: " + ex.ToString();
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //    }
//            //    /* Save end time and close the Excel object */
//            //    xls_excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//            //    xls_excelWorkBookRouterIntegrationTest.Save();
//            //    /* Close excel application when finish one job(power level) */
//            //    closeExcelRouterIntegrationTest();
//            //    Thread.Sleep(3000);
//            //    //寄出報告
//            //    try
//            //    {
//            //        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Read Excel File Failed", "Read Excel File Failed!!!", savePath);
//            //    }
//            //    catch (Exception ex)
//            //    {
//            //        str = "Send Mail Failed: " + ex.ToString();
//            //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //    }
//            //    return false;
//            //}

//            //str = " Succeed!!";
//            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            /* Run Main condition for Excel Data */
//            for (int rowSanity = 0; rowSanity < testConfigRouterIntegrationSanityTest.GetLength(0); rowSanity++)
//            {                
//                int iIndexRead;
//                string iIndex = string.Empty;
//                string sFunctionName = string.Empty;
//                string sName = string.Empty;
//                bool bHasInfo = false;
//                s_RouterIntegrationFinalReportInfo = "";

//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    bRouterIntegrationTestFTThreadRunning = false;
//                    //MessageBox.Show("Abort test", "Error");
//                    //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                    //threadRouterIntegrationTestFT.Abort();
//                    return false;
//                    // Never go here ...
//                }

//                if (testConfigRouterIntegrationSanityTest[rowSanity, 0] == "" || testConfigRouterIntegrationSanityTest[rowSanity, 0] == null)
//                {                    
//                    continue;
//                }

//                if (!Int32.TryParse(testConfigRouterIntegrationSanityTest[rowSanity, 0].Trim(), out iIndexRead))
//                {                    
//                    continue;
//                }

//                iIndex = testConfigRouterIntegrationSanityTest[rowSanity, 0];
//                sFunctionName = testConfigRouterIntegrationSanityTest[rowSanity, 1];
//                sName = testConfigRouterIntegrationSanityTest[rowSanity, 2];                
                
//                str = "====Index " + iIndex + " Start=====";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                i_RouterIntegrationCurrentIndex = iIndexRead;
//                string sLogFile = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_RouterIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";

//                str = String.Format("Run condition: Function Name: {0}, Name: {1}", sFunctionName, sName);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                
//                try
//                {
//                    /* Fill Loop, PowerLevel and Constellation in Excel */
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = iIndex;
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3] = sFunctionName;
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 4] = sName;                    
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine(ex);
//                }

//                /* Run each Sanity Test Item */
//                int iStartIndex = -1;
//                int iStopIndex = -1;

//                if(!RouterIntegrationTestGetSanityIndex(sFunctionName, ref iStartIndex, ref iStopIndex))
//                {
//                    try
//                    {
//                        /* Fill Loop, PowerLevel and Constellation in Excel */
//                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5] = "Error";
//                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 6] = "Function Unknown or Not Support";
//                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7] = sLogFile;
                        
//                        xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

//                        /* Write Console Log File as HyperLink in Excel report */
//                        xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7];
//                        xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sLogFile, Type.Missing, "ConsoleLog", "ConsoleLog");            
//                    }
//                    catch (Exception ex)
//                    {
//                        Debug.WriteLine(ex);
//                    }

//                    str = "Function Unknown or Not Support!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
                    
//                    m_PositionRouterIntegrationTest++;
//                    continue;
//                }

//                //switch (sFunctionName)
//                //{
//                //    case "Verify DUT can be reset by GUI":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //        break;
//                //    case "Remote factory reset via SNMP":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_RemoteFactoryResetViaSnmpStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_RemoteFactoryResetViaSnmpStopIndex;
//                //        break;
//                //    case "Upstream and downstream can be locked":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_UpstreamAndDownstreamCanBeLockedStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_UpstreamAndDownstreamCanBeLockedStopIndex;
//                //        break;
//                //    case "MDD in Docsis2.0/3.0 mode MUST be work":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_MddInDocsis20_30ModeMustBeWorkStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_MddInDocsis20_30ModeMustBeWorkStopIndex;
//                //        break;
//                //    case "tchCmAPBpi2CertStatus":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_tchCmAPBpi2CertStatusTestStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_tchCmAPBpi2CertStatusTestStopIndex;
//                //        break;
//                //    case "tchVendorDefaultDSfreq":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_tchVendorDefaultDSfreqTestStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_tchVendorDefaultDSfreqTestStopIndex;
//                //        break;
//                //    case "tchCmForceDualscan":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_tchCmForceDualscanTestStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_tchCmForceDualscanTestStopIndex;
//                //        break;
//                //    case "SNR & DS/US Power":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_SnrAndDsUsPowerStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_SnrAndDsUsPowerStopIndex;
//                //        break;
//                //    case "CM SNMP Agent":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_CmSnmpAgentStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_CmSnmpAgentStopIndex;
//                //        break;
//                //    case "MTA SNMP Agent":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_MtaSnmpAgentStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_MtaSnmpAgentStopIndex;
//                //        break;
//                //    case "SSID and Password generate rule":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_SsidAndPasswordGenerateRuleStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_SsidAndPasswordGenerateRuleStopIndex;
//                //        break;
//                //    case "5 GHz and 2.4 GHz Each band can be selected":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_5g24gEachBandCanBeSelectedStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_5g24gEachBandCanBeSelectedStopIndex;
//                //        break;
//                //    //case "":
//                //    //    iStartIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;

//                //    //    case "":
//                //    //    iStartIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;
//                //    //    case "":
//                //    //    iStartIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;
//                //    //    case "":
//                //    //    iStartIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;
//                //    //    case "":
//                //    //    iStartIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;
//                //    //    case "":
//                //    //    iStartIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStartIndex;
//                //    //    iStopIndex = i_RouterIntegrationSanityFunction_VerifyDutCanBeResetByGuiStopIndex;
//                //    //    break;
//                //    case "For Test":
//                //        iStartIndex = i_RouterIntegrationSanityFunction_ForTestStartIndex;
//                //        iStopIndex = i_RouterIntegrationSanityFunction_ForTestStopIndex;
//                //        break;
                    
//                //    default:                        
//                //        break;                        
//                //}

//                //str = String.Format("Sub-Index: Start:{0}, Stop: {1}", iStartIndex.ToString(), iStopIndex.ToString());
//                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //if(iStartIndex == -1 || iStopIndex == -1)
//                //{                   
//                //    try
//                //    {
//                //        /* Fill Loop, PowerLevel and Constellation in Excel */
//                //        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5] = "Error";
//                //        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 6] = "Function Unknown or Not Support";
//                //        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7] = sLogFile;
                        
//                //        xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

//                //        /* Write Console Log File as HyperLink in Excel report */
//                //        xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7];
//                //        xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sLogFile, Type.Missing, "ConsoleLog", "ConsoleLog");            
//                //    }
//                //    catch (Exception ex)
//                //    {
//                //        Debug.WriteLine(ex);
//                //    }

//                //    str = "Function Unknown or Not Support!!";
//                //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //    File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
//                //    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });
                    
//                //    m_PositionRouterIntegrationTest++;
//                //    continue;
//                //}

//                s_RouterIntegrationFinalReportInfo = "";
//                bHasInfo = false;

//                /* Run Condition IndexStart to IndexStop */
//                for (int index = iStartIndex; index <= iStopIndex; index++)
//                {
//                    iIndexCount = 0;
//                    sa_TestConditionSets = null;
//                    sa_ReportRouterIntegration = null;
//                    GC.Collect();

//                    if (bRouterIntegrationTestFTThreadRunning == false)
//                    {
//                        bRouterIntegrationTestFTThreadRunning = false;                        
//                        return false;
//                        // Never go here ...
//                    }

//                    str = "== Run sub-Index " + index.ToString() + " Start==";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    for (int row = 0; row < testConfigRouterIntegrationTest.GetLength(0); row++)
//                    {                        
                        
                        
//                        //i_RouterIntegrationCurrentIndex = index;
//                        sa_ReportRouterIntegration = new string[i_RouterIntegrationConditionParamterNum + 3];

//                        if (bRouterIntegrationTestFTThreadRunning == false)
//                        {
//                            bRouterIntegrationTestFTThreadRunning = false;
//                            return false;
//                            // Never go here ...
//                        }

//                        if (testConfigRouterIntegrationTest[row, 0] == "" || testConfigRouterIntegrationTest[row, 0] == null)
//                        {
//                            continue;
//                        }

//                        if (!Int32.TryParse(testConfigRouterIntegrationTest[row, 0].Trim(), out iIndexRead))
//                        {
//                            continue;
//                        }

//                        if (iIndexRead == index)
//                        {
//                            if (sa_TestConditionSets == null)
//                            {
//                                //sa_TestConditionSets = new string[RouterIntegrationDutDataSets.Length, i_RouterIntegrationConditionParamterNum];
//                                sa_TestConditionSets = new string[30, i_RouterIntegrationConditionParamterNum];
//                            }

//                            for (int j = 0; j < testConfigRouterIntegrationTest.GetLength(1); j++)
//                            {
//                                sa_TestConditionSets[iIndexCount, j] = testConfigRouterIntegrationTest[row, j];
//                            }

//                            iIndexCount++;
//                            iCount++;
//                        }
//                    } //End of for- loop: Scan every row of testConfigRouterIntegrationTest

//                    if (sa_TestConditionSets == null || sa_TestConditionSets[0, 1] == null)
//                    { // 找不到有此Index的欄位, 直接判定 Test Failed.
//                        try
//                        {
//                            /* Fill Loop, PowerLevel and Constellation in Excel */
//                            xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5] = "Fail";
//                            xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 6] = "Sub-Index:" +index.ToString() + "Not Existt";
//                            xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7] = sLogFile;

//                            xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

//                            /* Write Console Log File as HyperLink in Excel report */
//                            xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7];
//                            xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sLogFile, Type.Missing, "ConsoleLog", "ConsoleLog");
//                        }
//                        catch (Exception ex)
//                        {
//                            Debug.WriteLine(ex);
//                        }

//                        str = "Sub-Index:" +index.ToString() + "Not Existt";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                        File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                        m_PositionRouterIntegrationTest++;
//                        continue;
//                    }
                    
//                    str = "Current Device Index: " + RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].index.ToString() + ", SKU: " + RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = "Run Sub Function : " + sa_TestConditionSets[0, 1];
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    //string[] sReport = new string[i_RouterIntegrationConditionParamterNum + 3];

//                    /* Run Sub function */
//                    sa_ReportRouterIntegration[0] = sa_TestConditionSets[0, 0];
//                    sa_ReportRouterIntegration[1] = sa_TestConditionSets[0, 1];
//                    sa_ReportRouterIntegration[2] = sa_TestConditionSets[0, 2];

//                    bool bMultiReport = false;

//                    switch (sa_TestConditionSets[0, 1])
//                    {
//                        case "SNMP":
//                            RouterIntegrationSnmpFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "CMTS":
//                            RouterIntegrationCmtsFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "DUTCONSOLE":
//                            RouterIntegrationDutFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "WEB":
//                            RouterIntegrationGuiFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            bMultiReport = true;
//                            break;
//                        case "CNR_WEB":
//                            RouterIntegrationGuiFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            bMultiReport = true;
//                            break;
//                        case "SWITCH":
//                            RouterIntegrationSwtichFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "REMOTE":
//                            RouterIntegrationRemoteControlFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "ResetDevice":
//                            RouterIntegrationResetDeviceFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "Wait":
//                            RouterIntegrationWaitFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "Ftp":
//                            RouterIntegrationFtpFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "Ping":
//                            RouterIntegrationPingFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "EditFile":
//                            RouterIntegrationEditFileFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "CfgConvert":
//                            RouterIntegrationConvertCfgFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "DownloadFwFile":
//                            RouterIntegrationDownloadFwFileFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "CheckVer":
//                            RouterIntegrationCheckVersionFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;
//                        case "UploadFileToServer":
//                            RouterIntegrationUploadFileToSerrverFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                            break;


                            
                            
                            


                            
//                        //case "Ping":
//                        //    RouterIntegrationPingFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                        //    break;
//                        //case "SNMP":
//                        //    break;
//                        //case "SNMP":
//                        //    break;
//                        default:
//                            sa_ReportRouterIntegration[3] = "Error";
//                            sa_ReportRouterIntegration[4] = "Function Name Not Support!!";
//                            for (int reportIndex = 6; reportIndex < sa_ReportRouterIntegration.Length; reportIndex++)
//                            {
//                                if (sa_TestConditionSets[0, reportIndex - 3] != null)
//                                    sa_ReportRouterIntegration[reportIndex] = sa_TestConditionSets[0, reportIndex - 3];
//                            }
//                            break;
//                    }

//                    if (sa_ReportRouterIntegration[3] == "INFO")
//                    {
//                        bHasInfo = true;
//                    }

//                    if (sa_ReportRouterIntegration[3] == "Error")
//                    { //任何一步有錯誤, 一律中斷, 直接寫 Report, 跳下一個Sanity測試項
//                        //string sLogFile = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_RouterIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//                        sa_ReportRouterIntegration[5] = sLogFile;
//                        s_RouterIntegrationFinalReportInfo = sa_ReportRouterIntegration[4];

//                        RouterIntegrationTestReportData_Sanity(sa_ReportRouterIntegration);

//                        File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                        index = iStopIndex + 1;
//                        //m_PositionRouterIntegrationTest++;
//                        continue;
//                        //return false;
//                    }
//                    else if (sa_ReportRouterIntegration[3] == "FAIL")
//                    { //任何一步有錯誤, 一律中斷, 直接寫 Report, 跳下一個Sanity測試項
//                        //string sLogFile = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_RouterIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//                        sa_ReportRouterIntegration[5] = sLogFile;
//                        //s_RouterIntegrationFinalReportInfo = sa_ReportRouterIntegration[4];
//                        s_RouterIntegrationFinalReportInfo = "  FAIL in : " + sa_TestConditionSets[0, 2];
//                        //if (s_RouterIntegrationFinalReportInfo == "")
//                        //{
//                        //    s_RouterIntegrationFinalReportInfo = "  FAIL in : " + sa_TestConditionSets[0, 2];
//                        //}
//                        //else
//                        //{
//                        //    s_RouterIntegrationFinalReportInfo += sa_ReportRouterIntegration[4];
//                        //}

//                        RouterIntegrationTestReportData_Sanity(sa_ReportRouterIntegration);

//                        File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                        index = iStopIndex + 1;
//                        //m_PositionRouterIntegrationTest++;
//                        continue;
//                        //return false;
//                    }
//                    else if (sa_ReportRouterIntegration[3] == "FAILFILE")
//                    { //任何一步有錯誤, 一律中斷, 直接寫 Report, 跳下一個Sanity測試項
//                        //string sLogFile = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_RouterIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//                        sa_ReportRouterIntegration[5] = sLogFile;
//                        //s_RouterIntegrationFinalReportInfo = sa_ReportRouterIntegration[4];

//                        s_RouterIntegrationFinalReportInfo = "  FAIL in : " + sa_TestConditionSets[0, 2] + sa_ReportRouterIntegration[4];

//                        //if (s_RouterIntegrationFinalReportInfo == "")
//                        //{
//                        //    s_RouterIntegrationFinalReportInfo = "  FAIL in : " + sa_TestConditionSets[0, 2];
//                        //}
//                        //else
//                        //{
//                        //    s_RouterIntegrationFinalReportInfo += sa_ReportRouterIntegration[4];
//                        //}

//                        RouterIntegrationTestReportData_Sanity(sa_ReportRouterIntegration);

//                        File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                        index = iStopIndex + 1;
//                        //m_PositionRouterIntegrationTest++;
//                        continue;
//                        //return false;
//                    }
//                    else
//                    {//PASS
//                        sa_ReportRouterIntegration[5] = sLogFile;

//                        //if (sa_ReportRouterIntegration[3] == "FILE" || sa_ReportRouterIntegration[3] == "PASSFILE")
//                        if (sa_ReportRouterIntegration[3] == "FILE" || sa_ReportRouterIntegration[3] == "PASSFILE")
//                        {                       
//                            RouterIntegrationTestReportData_Sanity(sa_ReportRouterIntegration);
//                            sa_ReportRouterIntegration[4] = "";
//                        }
//                        s_RouterIntegrationFinalReportInfo += sa_ReportRouterIntegration[4];

//                        //測試完成, 寫入最後結果
//                        //string sLogFile = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_RouterIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//                        //sa_ReportRouterIntegration[5] = sLogFile;
//                        // 若中間有INFO, Test Result 留白
//                        if (bHasInfo) sa_ReportRouterIntegration[3] = "";
//                        if (sa_ReportRouterIntegration[3] == "PASSFILE") sa_ReportRouterIntegration[3] = "PASS";

//                        RouterIntegrationTestReportData_Sanity(sa_ReportRouterIntegration);

//                        //File.WriteAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
//                        File.AppendAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
//                        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                        //m_PositionRouterIntegrationTest++;
//                    }                   

//                } //End of for-loop: Run IndexStart to IndexStop 
//                m_PositionRouterIntegrationTest++;
//            }
           
//            /* Save end time and close the Excel object */
//            xls_excelWorkSheetRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//            saveExcelRouterIntegrationTest(savePath);
//            //xls_excelWorkBookRouterIntegrationTest.Save();
//            /* Close excel application when finish one job(power level) */
//            closeExcelRouterIntegrationTest();
//            Thread.Sleep(3000);

//            try
//            {
//                SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : " + RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel + " Test Report", "Test Completed!!!", savePath);
//            }
//            catch (Exception ex)
//            {
//                str = "Send Mail Failed: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            }

//            File.AppendAllText(saveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//            Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//            return true;


//            ///* Save end time and close the Excel object */
//            //xls_excelWorkSheetRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//            //saveExcelRouterIntegrationTest(savePath);
//            ////xls_excelWorkBookRouterIntegrationTest.Save();
//            ///* Close excel application when finish one job(power level) */
//            //closeExcelRouterIntegrationTest();
//            //Thread.Sleep(3000);

//            //try
//            //{
//            //    SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : " + RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel + " Test Report", "Test Completed!!!", savePath);
//            //}
//            //catch (Exception ex)
//            //{
//            //    str = "Send Mail Failed: " + ex.ToString();
//            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });                     
//            //}
            
//            //File.WriteAllText(saveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//            //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//            //return true;
//        }

//        //private bool RouterIntegrationTestMainFunction_Sanity2(int iStartIndex, int iStopIndex)
//        //{

//        //}
        
//        /* For reference , This function want't be used */
//        private bool RouterIntegrationTestMainFunction(string sExcelFileName)
//        {
//            string str = string.Empty;
//            int iCount = 0; // Indicate how many excel index row has read.
//            int iIndexCount = 0; // Indicate same index has read.

//            int iIndexCurrentRun = 1;
//            string sExcelFile = string.Empty;

//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();

//            /* Initial Value */
//            m_PositionRouterIntegrationTest = 16;
//            testConfigRouterIntegrationTest = null;

//            string subFolder = m_RouterIntegrationTestSubFolder + "\\" + RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel;
//            string savePath = createRouterIntegrationTestSaveExcelFile(subFolder, RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel, m_LoopRouterIntegrationTest, sExcelFileName);
//            string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//            //string savePath = createRouterIntegrationTestSaveExcelFile(m_RouterIntegrationTestSubFolder, RouterIntegrationDutDataSets[i].SkuModel, m_LoopRouterIntegrationTest, sExcelFileName);
//            /* initial Excel component */
//            initialExcelRouterIntegrationTest(savePath);

//            try
//            {
//                /* Fill Loop, PowerLevel and Constellation in Excel */
//                xls_excelWorkSheetRouterIntegrationTest.Cells[9, 3] = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SerialNumber;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[10, 3] = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SwVersion;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[11, 3] = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].HwVersion;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[4, 9] = m_LoopRouterIntegrationTest;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[5, 9] = Path.GetFileName(sExcelFileName);
//                xls_excelWorkSheetRouterIntegrationTest.Cells[6, 9] = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[7, 9] = "";
//                xls_excelWorkSheetRouterIntegrationTest.Cells[8, 9] = "";
//            }
//            catch (Exception ex)
//            {
//                Debug.WriteLine(ex);
//            }

//            /* Read excel content of excel file */
//            if (!ReadTestConfig_ReadExcelRouterIntegrationTest(sExcelFileName))
//            {
//                str = " Failed!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                //string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//                File.WriteAllText(saveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//                Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                try
//                {
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = "Read File Failed";

//                    /* Write Console Log File as HyperLink in Excel report */
//                    xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3];
//                    xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//                }
//                catch (Exception ex)
//                {
//                    str = "Write Data to Excel File Exception: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                }
//                /* Save end time and close the Excel object */
//                xls_excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                xls_excelWorkBookRouterIntegrationTest.Save();
//                /* Close excel application when finish one job(power level) */
//                closeExcelRouterIntegrationTest();
//                Thread.Sleep(3000);
//                //寄出報告
//                try
//                {
//                    SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Read Excel File Failed", "Read Excel File Failed!!!", savePath);
//                }
//                catch (Exception ex)
//                {
//                    str = "Send Mail Failed: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                }
//                return false;
//            }

//            str = " Succeed!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            /* Run Main condition for Excel Data */
//            for (int index = 1; index <= i_RouterIntegrationTestIndexMaximun; index++)
//            {
//                iIndexCount = 0;
//                sa_TestConditionSets = null;
//                sa_ReportRouterIntegration = null;
//                GC.Collect();
//                i_RouterIntegrationCurrentIndex = index;
//                sa_ReportRouterIntegration = new string[i_RouterIntegrationConditionParamterNum + 3];


//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    bRouterIntegrationTestFTThreadRunning = false;
//                    //MessageBox.Show("Abort test", "Error");
//                    //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                    //threadRouterIntegrationTestFT.Abort();
//                    return false;
//                    // Never go here ...
//                }

//                /* Check if all the excel content has tested. */
//                int iTotalCount = testConfigRouterIntegrationTest.GetLength(0);
//                if (iCount >= testConfigRouterIntegrationTest.GetLength(0)) break;

//                str = "====Index " + index.ToString() + " Start=====";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                for (int row = 0; row < testConfigRouterIntegrationTest.GetLength(0); row++)
//                {
//                    int iIndexRead;

//                    if (bRouterIntegrationTestFTThreadRunning == false)
//                    {
//                        bRouterIntegrationTestFTThreadRunning = false;
//                        //MessageBox.Show("Abort test", "Error");
//                        //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                        //threadRouterIntegrationTestFT.Abort();
//                        return false;
//                        // Never go here ...
//                    }

//                    if (testConfigRouterIntegrationTest[row, 0] == "" || testConfigRouterIntegrationTest[row, 0] == null)
//                    {
//                        if (index == 1)
//                        { //Count only once
//                            iCount++;
//                        }
//                        continue;
//                    }

//                    if (!Int32.TryParse(testConfigRouterIntegrationTest[row, 0].Trim(), out iIndexRead))
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
//                            //sa_TestConditionSets = new string[RouterIntegrationDutDataSets.Length, i_RouterIntegrationConditionParamterNum];
//                            sa_TestConditionSets = new string[30, i_RouterIntegrationConditionParamterNum];
//                        }

//                        for (int j = 0; j < testConfigRouterIntegrationTest.GetLength(1); j++)
//                        {
//                            sa_TestConditionSets[iIndexCount, j] = testConfigRouterIntegrationTest[row, j];
//                        }

//                        iIndexCount++;
//                        iCount++;
//                    }
//                } //End of excel content reading for-loop

//                if (sa_TestConditionSets == null || sa_TestConditionSets[0, 1] == null) continue;

//                str = "Current Device Index: " + RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].index.ToString() + ", SKU: " + RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = "Run Sub Function : " + sa_TestConditionSets[0, 1];
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                //string[] sReport = new string[i_RouterIntegrationConditionParamterNum + 3];

//                /* Run Sub function */
//                sa_ReportRouterIntegration[0] = sa_TestConditionSets[0, 0];
//                sa_ReportRouterIntegration[1] = sa_TestConditionSets[0, 1];
//                sa_ReportRouterIntegration[2] = sa_TestConditionSets[0, 2];

//                switch (sa_TestConditionSets[0, 1])
//                {
//                    case "SNMP":
//                        RouterIntegrationSnmpFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                        break;
//                    case "CMTS":
//                        RouterIntegrationCmtsFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                        break;
//                    case "DUTCONSOLE":
//                        RouterIntegrationDutFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                        break;
//                    case "WEB":
//                        //RouterIntegrationSnmpFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                        break;
//                    case "CNRWEB":
//                        //RouterIntegrationSnmpFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                        break;                        
//                    case "SWITCH":
//                        RouterIntegrationSwtichFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                        break;                        
//                    case "REMOTETCP":
//                        //RouterIntegrationSnmpFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                        break;
//                    //case "SNMP":
//                    //    break;
//                    //case "SNMP":
//                    //    break;
//                    default:
//                        sa_ReportRouterIntegration[3] = "Error";
//                        sa_ReportRouterIntegration[4] = "Function Name Not Support!!";
//                        for (int reportIndex = 6; reportIndex < sa_ReportRouterIntegration.Length; reportIndex++)
//                        {
//                            if (sa_TestConditionSets[0, reportIndex - 3] != null)
//                                sa_ReportRouterIntegration[reportIndex] = sa_TestConditionSets[0, reportIndex - 3];
//                        }
//                        break;
//                }

//                string sLogFile = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_RouterIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//                sa_ReportRouterIntegration[5] = sLogFile;

//                RouterIntegrationTestReportData(sa_ReportRouterIntegration);

//                File.WriteAllText(sLogFile, txtRouterIntegrationFunctionTestInformation.Text);
//                Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                //if (threadRouterIntegrationTestRunCondition != null)
//                //    threadRouterIntegrationTestRunCondition.Abort();

//                m_PositionRouterIntegrationTest++;
//            }




//            ///* Wait Time to process next mib */
//            //str = " Wait for Test Condition Finished or Timeout(S):" + nudRouterIntegrationFunctionTestConditionTimeout.Value.ToString();
//            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //bConditionStatus = false;

//            //Thread threadRouterIntegrationTestRunCondition = new Thread(DoRouterIntegrationTestRunCondition);
//            //threadRouterIntegrationTestRunCondition.Name = "RouterIntegrationTestRunCondition";
//            //threadRouterIntegrationTestRunCondition.Start();

//            //try
//            //{
//            //    stopWatch.Stop();
//            //    stopWatch.Reset();
//            //    stopWatch.Start();

//            //    while (true)
//            //    {
//            //        if (bRouterIntegrationTestFTThreadRunning == false)
//            //        {
//            //            bRouterIntegrationTestFTThreadRunning = false;
//            //            //MessageBox.Show("Abort test", "Error");
//            //            //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//            //            //threadRouterIntegrationTestFT.Abort();
//            //            return false;
//            //            // Never go here ...
//            //        }

//            //        if (bConditionStatus)
//            //            break;

//            //        if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestConditionTimeout.Value * 1000))
//            //        {
//            //            if (!bConditionStatus)
//            //            {
//            //                str = "Run Condition Timeout!!";
//            //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //                //if (threadRouterIntegrationTestRunCondition != null)
//            //                //    threadRouterIntegrationTestRunCondition.Abort();
//            //            }
//            //            break;
//            //        }
//            //    }
//            //}
//            //catch (Exception ex)
//            //{
//            //    str = "Exception: Run Condition Timeout: " + ex.ToString();
//            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //}

//            //if (!bConditionStatus)
//            //{ // Report Condition Timeout
//            //    for (int reportIndex = 0; reportIndex < 3; reportIndex++)
//            //    {
//            //        sa_ReportRouterIntegration[reportIndex] = sa_TestConditionSets[0, reportIndex];
//            //    }

//            //    sa_ReportRouterIntegration[3] = "Error";
//            //    sa_ReportRouterIntegration[4] = "Run Condition Timeout";

//            //    for (int reportIndex = 6; reportIndex < sa_ReportRouterIntegration.Length; reportIndex++)
//            //    {
//            //        if (sa_TestConditionSets[0, reportIndex - 3] != null)
//            //            sa_ReportRouterIntegration[reportIndex] = sa_TestConditionSets[0, reportIndex - 3];
//            //    }

//            //    //RouterIntegrationTestReportData(sReport);
//            //}

//            //string sLogFile = s_CableIntegrationSnmpsaveFileLog = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_RouterIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";
//            //sa_ReportRouterIntegration[5] = sLogFile;

//            //RouterIntegrationTestReportData(sa_ReportRouterIntegration);

//            //File.WriteAllText(s_CableIntegrationSnmpsaveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//            //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//            ////if (threadRouterIntegrationTestRunCondition != null)
//            ////    threadRouterIntegrationTestRunCondition.Abort();

//            //    m_PositionRouterIntegrationTest++;
//            //}
//            /* Save end time and close the Excel object */
//            xls_excelWorkSheetRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//            saveExcelRouterIntegrationTest(savePath);
//            //xls_excelWorkBookRouterIntegrationTest.Save();
//            /* Close excel application when finish one job(power level) */
//            closeExcelRouterIntegrationTest();
//            Thread.Sleep(3000);

//            try
//            {
//                SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : " + RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].SkuModel + " Test Report", "Test Completed!!!", savePath);
//            }
//            catch (Exception ex)
//            {
//                str = "Send Mail Failed: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            }

//            File.WriteAllText(saveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//            Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//            return true;
//        }

//        /* For reference , This function want't be used */
//        private bool RouterIntegrationTestMainFunctionOrig(string sExcelFileName)
//        {
//            string str = string.Empty;
//            int iCount = 0; // Indicate how many excel index row has read.
//            int iIndexCount = 0; // Indicate same index has read.
            
//            int iIndexCurrentRun = 1;
//            string sExcelFile = string.Empty;

//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();            

//            /* Initial Value */
//            m_PositionRouterIntegrationTest = 16;
//            testConfigRouterIntegrationTest = null;          

//            /* Start to run main function for each Data Sets */
//            for (int i = 0; i < RouterIntegrationDutDataSets.Length; i++)
//            {
//                i_RouterIntegrationCurrentDateSetsIndex = i;
//                iCount = 0;
//                m_PositionRouterIntegrationTest = 16;
//                string saveLogSubPath = m_RouterIntegrationTestSubFolder + "\\" + RouterIntegrationDutDataSets[i].SkuModel + "\\" + "HeaderLog.txt";

//                if (RouterIntegrationDutDataSets[i].TestType.ToLower().IndexOf("throughput") >= 0)
//                {
//                    RouterIntegrationVeriwaveMainFunction(i_RouterIntegrationCurrentDateSetsIndex);
//                    continue;
//                }

//                /* Set Comport to device setting comport */
//                str = "Set Comport to : " + RouterIntegrationDutDataSets[i].ComportNum;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                comPortRouterIntegrationTest.Close();
//                comPortRouterIntegrationTest.SetPortName(RouterIntegrationDutDataSets[i].ComportNum);
//                comPortRouterIntegrationTest.Open();

//                if (comPortRouterIntegrationTest.isOpen() != true)
//                {
//                    //MessageBox.Show("COM port is not ready!");
//                    //System.Windows.Forms.Cursor.Current = Cursors.Default;
//                    //btnRouterIntegrationFunctionTestRun.Enabled = true;
//                    str = "Set Comport Failed : " + RouterIntegrationDutDataSets[i].ComportNum;
//                    File.AppendAllText(saveLogSubPath, txtRouterIntegrationFunctionTestInformation.Text);
                    
//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Download FW Process Failed", "Copy Rd FW to NFS server Process Failed!!!", saveLogSubPath);
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Send Mail Failed: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }

//                    continue;
//                }

//                str = "Done.";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });                

//                //if (!CGA2121CheckFwFileName_RouterIntegrationTest())
//                //{ //Read show cfg failed, or the FW filename is wrong 

//                //    str = "Show version failed, or the FW filename is wrong";
//                //    Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                //    File.AppendAllText(saveLogSubPath, txtRouterIntegrationFunctionTestInformation.Text);
                    
//                //    //寄出報告
//                //    try
//                //    {
//                //        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Download FW Process Failed", "Copy Rd FW to NFS server Process Failed!!!", saveLogSubPath);
//                //    }
//                //    catch (Exception ex)
//                //    {
//                //        str = "Send Mail Failed: " + ex.ToString();
//                //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                //    }

//                //    continue;
//                //}

//                /* Create excel report and conosle log path */
//                /* Format : string PathFile = @"\report\SUBFOLDER\FileName_Report_DATE_(LOOP).xlsx";*/
//                string subFolder = m_RouterIntegrationTestSubFolder + "\\" + RouterIntegrationDutDataSets[i].SkuModel;
//                string savePath = createRouterIntegrationTestSaveExcelFile(subFolder, RouterIntegrationDutDataSets[i].SkuModel, m_LoopRouterIntegrationTest, sExcelFileName);
//                //string savePath = createRouterIntegrationTestSaveExcelFile(m_RouterIntegrationTestSubFolder, RouterIntegrationDutDataSets[i].SkuModel, m_LoopRouterIntegrationTest, sExcelFileName);
//                /* initial Excel component */
//                initialExcelRouterIntegrationTest(savePath);

//                try
//                {
//                    /* Fill Loop, PowerLevel and Constellation in Excel */
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[4, 9] = m_LoopRouterIntegrationTest;
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[5, 9] = Path.GetFileName(sExcelFileName);
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[6, 9] = RouterIntegrationDutDataSets[i].SkuModel;
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[7, 9] = "";
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[8, 9] = "";
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine(ex);
//                }

//                /* Ping Device ip first */
//                str = "Ping Device IP Address: " + RouterIntegrationDutDataSets[i].IpAddress;
//                Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                if (!QuickPingRouterIntegrationTest(RouterIntegrationDutDataSets[i].IpAddress, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//                {
//                    str = " Failed!! Run next test condition...";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//                    File.WriteAllText(saveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                    try
//                    {
//                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = "Ping Failed";

//                        /* Write Console Log File as HyperLink in Excel report */
//                        xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3];
//                        xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Write Data to Excel File Exception: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }

//                    /* Save end time and close the Excel object */
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                    xls_excelWorkBookRouterIntegrationTest.Save();
//                    /* Close excel application when finish one job(power level) */
//                    closeExcelRouterIntegrationTest();
//                    Thread.Sleep(3000);

//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Ping Device Failed", "Ping Device Failed!!!", savePath);
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Send Mail Failed: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }

//                    //return false;
//                    continue;
//                }

//                str = " Succeed.";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = "Read Excel Content: " + sExcelFileName;
//                Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                /* Read excel content of excel file */
//                if (!ReadTestConfig_ReadExcelRouterIntegrationTest(sExcelFileName))
//                {
//                    str = " Failed!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//                    File.WriteAllText(saveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//                    Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                    try
//                    {
//                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = "Read File Failed";

//                        /* Write Console Log File as HyperLink in Excel report */
//                        xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3];
//                        xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Write Data to Excel File Exception: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }
//                    /* Save end time and close the Excel object */
//                    xls_excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                    xls_excelWorkBookRouterIntegrationTest.Save();
//                    /* Close excel application when finish one job(power level) */
//                    closeExcelRouterIntegrationTest();
//                    Thread.Sleep(3000);

//                    //寄出報告
//                    try
//                    {
//                        SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Read Excel File Failed", "Read Excel File Failed!!!", savePath);
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Send Mail Failed: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    }
//                    continue;
//                    //return false;
//                }

//                str = " Succeed!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                ///* Open Comport of Datasets */
//                //comPortRouterIntegrationTest.ResetPort(RouterIntegrationDutDataSets[i].ComportNum);
//                //if (comPortRouterIntegrationTest.isOpen() != true)
//                //{
//                //    //MessageBox.Show("COM port is not ready!");
//                //    //System.Windows.Forms.Cursor.Current = Cursors.Default;
//                //    //btnRouterIntegrationFunctionTestRun.Enabled = true;
//                //    //return;
//                //}

//                /* Run Main condition for Excel Data */
//                for (int index = 1; index <= i_RouterIntegrationTestIndexMaximun; index++)
//                {
//                    iIndexCount = 0;
//                    sa_TestConditionSets = null;
//                    i_RouterIntegrationCurrentIndex = index;                    

//                    if (bRouterIntegrationTestFTThreadRunning == false)
//                    {
//                        bRouterIntegrationTestFTThreadRunning = false;
//                        //MessageBox.Show("Abort test", "Error");
//                        //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                        //threadRouterIntegrationTestFT.Abort();
//                        return false;
//                        // Never go here ...
//                    }

//                    /* Check if all the excel content has tested. */
//                    int iTotalCount = testConfigRouterIntegrationTest.GetLength(0);
//                    if (iCount >= testConfigRouterIntegrationTest.GetLength(0)) break;

//                    str = "====Index " + index.ToString() + " Start=====";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    for (int row = 0; row < testConfigRouterIntegrationTest.GetLength(0); row++)
//                    {                        
//                        int iIndexRead;

//                        if (bRouterIntegrationTestFTThreadRunning == false)
//                        {
//                            bRouterIntegrationTestFTThreadRunning = false;
//                            //MessageBox.Show("Abort test", "Error");
//                            //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                            //threadRouterIntegrationTestFT.Abort();
//                            return false;
//                            // Never go here ...
//                        }

//                        if (testConfigRouterIntegrationTest[row, 0] == "" || testConfigRouterIntegrationTest[row, 0] == null)
//                        {
//                            if (index == 1)
//                            { //Count only once
//                                iCount++;
//                            }
//                            continue;
//                        }

//                        if (!Int32.TryParse(testConfigRouterIntegrationTest[row, 0].Trim(), out iIndexRead))
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
//                                //sa_TestConditionSets = new string[RouterIntegrationDutDataSets.Length, i_RouterIntegrationConditionParamterNum];
//                                sa_TestConditionSets = new string[100, i_RouterIntegrationConditionParamterNum];
//                            }

//                            for (int j = 0; j < testConfigRouterIntegrationTest.GetLength(1); j++)
//                            {
//                                sa_TestConditionSets[iIndexCount, j] = testConfigRouterIntegrationTest[row, j];
//                            }

//                            iIndexCount++;
//                            iCount++;
//                        }
//                    } //End of excel content reading for-loop

//                    if (sa_TestConditionSets == null  || sa_TestConditionSets[0, 1] == null) continue;

//                    str = "Current Device Index: " + RouterIntegrationDutDataSets[i].index.ToString() + ", SKU: " + RouterIntegrationDutDataSets[i].SkuModel;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = "Run Sub Function : " + sa_TestConditionSets[0, 1];
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    string[] sReport = new string[i_RouterIntegrationConditionParamterNum + 3];
                    
//                    /* Wait Time to process next mib */
//                    str = " Wait for Test Condition Finished or Timeout(S):" + nudRouterIntegrationFunctionTestConditionTimeout.Value.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                    
//                    bConditionStatus = false;
                    
//                    try
//                    {
//                        Thread threadRouterIntegrationTestRunCondition = new Thread(DoRouterIntegrationTestRunCondition);
//                        threadRouterIntegrationTestRunCondition.Name = "RouterIntegrationTestRunCondition";
//                        threadRouterIntegrationTestRunCondition.Start();

//                        stopWatch.Stop();
//                        stopWatch.Reset();
//                        stopWatch.Start();

//                        while (true)
//                        {
//                            if (bRouterIntegrationTestFTThreadRunning == false)
//                            {
//                                bRouterIntegrationTestFTThreadRunning = false;
//                                //MessageBox.Show("Abort test", "Error");
//                                //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                                //threadRouterIntegrationTestFT.Abort();
//                                return false;
//                                // Never go here ...
//                            }

//                            if (bConditionStatus)
//                                break;

//                            if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestConditionTimeout.Value * 1000))
//                            {
//                                if (!bConditionStatus)
//                                {
//                                    str = "Run Condition Timeout!!";
//                                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                                    if (threadRouterIntegrationTestRunCondition != null)
//                                        threadRouterIntegrationTestRunCondition.Abort();
//                                }
//                                break;
//                            }
//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        str = "Exception: Run Condition Timeout: " + ex.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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

//                        RouterIntegrationTestReportData(sReport);
//                    }                    

//                    m_PositionRouterIntegrationTest++;
//                }
//                /* Save end time and close the Excel object */
//                xls_excelWorkSheetRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                saveExcelRouterIntegrationTest(savePath);
//                //xls_excelWorkBookRouterIntegrationTest.Save();
//                /* Close excel application when finish one job(power level) */
//                closeExcelRouterIntegrationTest();
//                Thread.Sleep(3000);
                               
//                try
//                {
//                    SendReportByGmailWithFile(s_EmailSenderID, s_EmailSenderPassword, s_EmailReceivers, "Cable SW Integration ATE Test : Test Report", "Test Completed!!!", savePath);
//                }
//                catch (Exception ex)
//                {
//                    str = "Send Mail Failed: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                }
//            }
           
//            return true;
//        }       
        
//        private void DoRouterIntegrationTestRunCondition()
//        {
//            //string[] sReport = new string[i_RouterIntegrationConditionParamterNum + 3];
//            //string sLogFile = s_CableIntegrationSnmpsaveFileLog = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].ReportSavePath + "\\ConsoleLog_Index" + i_RouterIntegrationCurrentIndex.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";

//            sa_ReportRouterIntegration[0] = sa_TestConditionSets[0, 0];
//            sa_ReportRouterIntegration[1] = sa_TestConditionSets[0, 1];
//            sa_ReportRouterIntegration[2] = sa_TestConditionSets[0, 2]; 

//            switch (sa_TestConditionSets[0, 1])
//            {

//                case "SNMP":
//                    RouterIntegrationSnmpFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);                    
//                    break;
//                case "CMTS":
//                    RouterIntegrationCmtsFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                    break;
//                case "DUTCONSOLE":
//                    RouterIntegrationDutFunction(ref sa_ReportRouterIntegration, i_RouterIntegrationCurrentIndex, i_RouterIntegrationCurrentDateSetsIndex);
//                    break;
//                case "WEB":
//                    break;
//                //case "SNMP":
//                //    break;
//                //case "SNMP":
//                //    break;
//                default:
//                    sa_ReportRouterIntegration[3] = "Error";
//                    sa_ReportRouterIntegration[4] = "Function Name Not Support!!";
//                    break;
//            }
            
//            //sReport[5] = sLogFile;

//            //RouterIntegrationTestReportData(sReport);

//            //File.WriteAllText(s_CableIntegrationSnmpsaveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//            //Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//            bConditionStatus = true;

//        }
    
//        //private bool CmReChanProcessRouterIntegrationTest()
//        //{            
//        //    string str = string.Empty;
//        //    //string sGoto_dsCheck = "Moving to Downstream Frequency 333000000 Hz";
//        //    string sGoto_dsCheck = "moving to downstream frequency " + nudRouterIntegrationFunctionTestGotoChannel.Value.ToString() + "000000 hz";
//        //    string sCableConfigCfgFile = System.Windows.Forms.Application.StartupPath + "\\importData\\" + txtRouterIntegrationFunctionTestCfgFileName.Text;

//        //    bool bCheckCrash = false;
//        //    bool bCrash = false;
//        //    bool bReChanProcess = false;
//        //    bool bCheckConfigFile = false;
//        //    bool bRightConfigFile = false;
//        //    string strCompare = string.Empty;
//        //    string sCheckConfigFile = string.Empty;
//        //    string NfsPath = txtRouterIntegrationConfigurationNfsLocation.Text;
//        //    b_Crash = false;

//        //    int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds

//        //    /* Go to root Directory */
//        //    if (!CGA2121BackToRootDirectory_RouterIntegrationTest("cm"))
//        //    {
//        //        return false;
//        //    }

//        //    // /* Re-Chan Device */
//        //    Stopwatch stopWatch = new Stopwatch();
//        //    stopWatch.Start();

//        //    if (iCmResetMode == 0)
//        //    {  //use dogo_ds freq command                

//        //        //comPortRouterIntegrationTest.DiscardBuffer();            
//        //        //Thread.Sleep(1000);

//        //        //comPortRouterIntegrationTest.WriteLine(" cd /");
//        //        //comPortRouterIntegrationTest.WriteLine("");
//        //        //str = comPortRouterIntegrationTest.ReadLine();
//        //        //if(str != "")
//        //        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //        //stopWatch.Stop();
//        //        //stopWatch.Reset();
//        //        //stopWatch.Restart();

//        //        //while (true)
//        //        //{
//        //        //    if (bRouterIntegrationTestFTThreadRunning == false)
//        //        //        return false;

//        //        //    if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//        //        //    {                    
//        //        //        str = "Change to directory: /CM Timeout." ;
//        //        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation }); 
//        //        //       return false;
//        //        //    }

//        //        //    if (str.ToLower().IndexOf("cm>") != -1) 
//        //        //    {
//        //        //        break;
//        //        //    }
//        //        //    else if (str.ToLower().IndexOf("rg>") != -1)
//        //        //    {
//        //        //        comPortRouterIntegrationTest.WriteLine("/Console/sw");
//        //        //        Thread.Sleep(3000);
//        //        //        comPortRouterIntegrationTest.WriteLine("");
//        //        //        str = comPortRouterIntegrationTest.ReadLine();
//        //        //        if(str != "")
//        //        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //        //    }

//        //        //    str = comPortRouterIntegrationTest.ReadLine();
//        //        //    if(str != "")
//        //        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });               
//        //        //}

//        //        str = "Run goto_ds " + nudRouterIntegrationFunctionTestGotoChannel.Value.ToString();
//        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //        str = "";

//        //        stopWatch.Stop();
//        //        stopWatch.Reset();
//        //        stopWatch.Restart();

//        //        long lTimeTemp = stopWatch.ElapsedMilliseconds + 30 * 1000;
//        //        comPortRouterIntegrationTest.WriteLine("/docsis_ctl/goto_ds " + nudRouterIntegrationFunctionTestGotoChannel.Value.ToString());
//        //        str = comPortRouterIntegrationTest.ReadLine();
//        //        if (str != "")
//        //        {
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //        }

//        //        while (true)
//        //        {
//        //            if (bRouterIntegrationTestFTThreadRunning == false)
//        //            {
//        //                bRouterIntegrationTestFTThreadRunning = false;
//        //                MessageBox.Show("Abort test", "Error");
//        //                this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //                threadRouterIntegrationTestFT.Abort();
//        //                // Never go here ...
//        //            }

//        //            if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestGotoWaitTime.Value) * 1000)
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
//        //                comPortRouterIntegrationTest.WriteLine("/docsis_ctl/goto_ds " + nudRouterIntegrationFunctionTestGotoChannel.Value.ToString());
//        //            }
//        //            str = comPortRouterIntegrationTest.ReadLine();
//        //            if (str != "")
//        //            {
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
//        //            if (bRouterIntegrationTestFTThreadRunning == false)
//        //            {
//        //                bRouterIntegrationTestFTThreadRunning = false;
//        //                MessageBox.Show("Abort test", "Error");
//        //                this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //                threadRouterIntegrationTestFT.Abort();
//        //                // Never go here ...
//        //            }

//        //            if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestGotoWaitTime.Value) * 1000)
//        //            {
//        //                break;
//        //            }

//        //            str = comPortRouterIntegrationTest.ReadLine();
//        //            if (str != "")
//        //            {
//        //                strCompare += str;
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //            File.Copy(s_CableConfigCfgBasicCfgFile, sCableConfigCfgFile, true);
//        //            CopyFileToNfsFoler(sCableConfigCfgFile, NfsPath, true);
//        //            str = "Copy File Succeed.";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
//        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //        str = "";

//        //        stopWatch.Stop();
//        //        stopWatch.Reset();
//        //        stopWatch.Restart();

//        //        long lTimeTemp = stopWatch.ElapsedMilliseconds + 30 * 1000;
//        //        comPortRouterIntegrationTest.WriteLine("/reset");
//        //        str = comPortRouterIntegrationTest.ReadLine();
//        //        if (str != "")
//        //        {
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //        }

//        //        while (true)
//        //        {
//        //            if (bRouterIntegrationTestFTThreadRunning == false)
//        //            {
//        //                bRouterIntegrationTestFTThreadRunning = false;
//        //                MessageBox.Show("Abort test", "Error");
//        //                this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //                threadRouterIntegrationTestFT.Abort();
//        //                // Never go here ...
//        //            }

//        //            if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestGotoWaitTime.Value) * 1000)
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
//        //                comPortRouterIntegrationTest.WriteLine("/reset");
//        //            }
//        //            str = comPortRouterIntegrationTest.ReadLine();
//        //            if (str != "")
//        //            {
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
//        //            if (bRouterIntegrationTestFTThreadRunning == false)
//        //            {
//        //                bRouterIntegrationTestFTThreadRunning = false;
//        //                MessageBox.Show("Abort test", "Error");
//        //                this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //                threadRouterIntegrationTestFT.Abort();
//        //                // Never go here ...
//        //            }

//        //            if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestGotoWaitTime.Value) * 1000)
//        //            {
//        //                break;
//        //            }

//        //            str = comPortRouterIntegrationTest.ReadLine();
//        //            if (str != "")
//        //            {
//        //                strCompare += str;
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                        
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
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //            File.Copy(s_CableConfigCfgBasicCfgFile, sCableConfigCfgFile, true);
//        //            CopyFileToNfsFoler(sCableConfigCfgFile, NfsPath, true);
//        //            str = "Copy File Succeed.";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
//        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //    //comPortRouterIntegrationTest.WriteLine(" cd /");
//        //    //comPortRouterIntegrationTest.WriteLine("");
//        //    //str = comPortRouterIntegrationTest.ReadLine();
//        //    //if (str != "")
//        //    //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //    //stopWatch.Stop();
//        //    //stopWatch.Reset();
//        //    //stopWatch.Restart();

//        //    //while (true)
//        //    //{
//        //    //    if (bRouterIntegrationTestFTThreadRunning == false)
//        //    //        return false;

//        //    //    if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//        //    //    {
//        //    //        str = "Change to directory: /CM Failed.";
//        //    //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //    //        return false;
//        //    //    }

//        //    //    if (str.ToLower().IndexOf("cm>") != -1)
//        //    //    {
//        //    //        break;
//        //    //    }
//        //    //    else if (str.ToLower().IndexOf("rg") != -1)
//        //    //    {
//        //    //        comPortRouterIntegrationTest.WriteLine("/Console/sw");
//        //    //        Thread.Sleep(3000);
//        //    //        comPortRouterIntegrationTest.WriteLine("");
//        //    //        str = comPortRouterIntegrationTest.ReadLine();
//        //    //        if (str != "")
//        //    //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //    //    }

//        //    //    str = comPortRouterIntegrationTest.ReadLine();
//        //    //    if (str != "")
//        //    //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //    //}

//        //    if (!CGA2121BackToRootDirectory_RouterIntegrationTest("cm"))
//        //    {
//        //        return false;
//        //    }

//        //    bool bCfgStatus = false;
//        //    s_CfgContect = "";
//        //    str = "";
//        //    comPortRouterIntegrationTest.DiscardBuffer();
//        //    Thread.Sleep(1000);
//        //    comPortRouterIntegrationTest.WriteLine("/Console/cm/show cfg");

//        //    str = comPortRouterIntegrationTest.ReadLine();
//        //    if (str != "")
//        //    {
//        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //        s_CfgContect += str;
//        //    }
      
//        //    stopWatch.Stop();
//        //    stopWatch.Reset();
//        //    stopWatch.Restart();

//        //    while (true)
//        //    {
//        //        if (bRouterIntegrationTestFTThreadRunning == false)
//        //        {
//        //            bRouterIntegrationTestFTThreadRunning = false;
//        //            MessageBox.Show("Abort test", "Error");
//        //            this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //            threadRouterIntegrationTestFT.Abort();
//        //            // Never go here ...
//        //        }

//        //        if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//        //        {
//        //            str = "Show cfg Timeout.";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //            s_CfgContect += str;    
//        //        }

//        //        str = comPortRouterIntegrationTest.ReadLine();
//        //        //Thread.Sleep(100);
//        //    } 

//        //    stopWatch.Stop();
//        //    stopWatch.Reset();

//        //    str = "Check Config File Name...";
//        //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //    /* Check if the device loaded cfg file file is equal to CableConfigFile.cfg */
//        //    string sConfigFile = Path.GetFileName(txtRouterIntegrationFunctionTestCfgFileName.Text);
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
//        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //        return false;               
//        //    }
//        //    return true;
            
//        //    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //    //RouterIntegrationTestReportData(ls_CurrentMibs, 1, bCrash, str);           
//        //}

//        private bool PingAndReadExcelFileRouterIntegrationTest(string savePath, string sExcelFileName, string ip)
//        {
//            string str = string.Empty;

//            /* Ping ip first */
//            if (s_TestType.ToLower() == "cm")
//            {
//                //str = "Ping CM IP Address: " + mtbRouterIntegrationFunctionTestCmIpAddress.Text;
//            }
//            else
//            {
//                //str = "Ping MTA IP Address: " + mtbRouterIntegrationFunctionTestMtaIpAddress.Text;
//            }
//            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
            
//            //if (!QuickPingRouterIntegrationTest(mtbRouterIntegrationFunctionTestMtaIpAddress.Text, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//            if (!QuickPingRouterIntegrationTest(ip, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//            {
//                str = " Failed!! Run next test condition...";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//                File.WriteAllText(saveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//                Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                try
//                {
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = "Ping Failed";

//                    /* Write Console Log File as HyperLink in Excel report */
//                    xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3];
//                    xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//                }
//                catch (Exception ex)
//                {
//                    str = "Write Data to Excel File Exception: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                }
//                /* Save end time and close the Excel object */
//                xls_excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                xls_excelWorkBookRouterIntegrationTest.Save();
//                /* Close excel application when finish one job(power level) */
//                closeExcelRouterIntegrationTest();
//                Thread.Sleep(3000);
//                return false;
//            }

//            str = " Succeed.";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            str = "Read Excel Content: " + sExcelFileName;
//            Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            /* Read excel content of excel file */
//            if (!ReadTestConfig_ReadExcelRouterIntegrationTest(sExcelFileName))
//            {
//                str = " Failed!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";
//                File.WriteAllText(saveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//                Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                try
//                {
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = "Read File Failed";

//                    /* Write Console Log File as HyperLink in Excel report */
//                    xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3];
//                    xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");
//                }
//                catch (Exception ex)
//                {
//                    str = "Write Data to Excel File Exception: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                }
//                /* Save end time and close the Excel object */
//                xls_excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                xls_excelWorkBookRouterIntegrationTest.Save();
//                /* Close excel application when finish one job(power level) */
//                closeExcelRouterIntegrationTest();
//                Thread.Sleep(3000);
//                return false;
//            }

//            str = " Succeed!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            return true;
//        }
        
//        //private bool RouterIntegrationTestMainFunction()
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
//        //    //    int snmpVersion = rbtnRouterIntegrationFunctionTestSnmpV1.Checked ? 1 : rbtnRouterIntegrationFunctionTestSnmpV2.Checked ? 2 : 3;
//        //    //    Snmp_RouterIntegrationTest = new CBTSnmp();
//        //    //    st_CbtMibDataConfigReadWriteTest = new CbtMibAccess[testConfigRouterIntegrationTest.GetLength(0)];
//        //    //    Snmp_RouterIntegrationTest.Init(mtbRouterIntegrationFunctionTestCmIpAddress.Text, snmpVersion, dudRouterIntegrationFunctionTestCmReadCommunity.Text, dudRouterIntegrationFunctionTestCmWriteCommunity.Text, Convert.ToInt32(nudRouterIntegrationFunctionTestSnmpPort.Value), Convert.ToInt32(nudRouterIntegrationFunctionTestTrapPort.Value));

//        //    //}
//        //    //catch (Exception ex)
//        //    //{
//        //    //    //MessageBox.Show(ex.ToString());
//        //    //    str = "SNMP objext Exception: " + ex.ToString();
//        //    //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //    //}      

//        //    ///
//        //    /// MIB Read Write Config test flow
//        //    ///
//        //    /* Main loop : Get test condition count for testing  */
//        //    for (int row = 0; row < testConfigRouterIntegrationTest.GetLength(0); row++)
//        //    {
//        //        if (bRouterIntegrationTestFTThreadRunning == false)
//        //        {
//        //            bRouterIntegrationTestFTThreadRunning = false;
//        //            MessageBox.Show("Abort test", "Error");
//        //            this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //            threadRouterIntegrationTestFT.Abort();
//        //            // Never go here ...
//        //        }

//        //        str = "====Index " +  indexCurrentSnmp.ToString() +" Start=====";
//        //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });


//        //        /* 因為比對的關係, 會造成最後一批無法被處理到, 因此多做這段處理, indexCurrentSnmp+1, 
//        //         * 可以跑else 內的部份, copy檔案及reset dut. 這也是因此row的condition 多加1的原因
//        //         * ,在read excel 中處理, row +1
//        //         * */
//        //        if (row == testConfigRouterIntegrationTest.GetLength(0) - 1)
//        //        {
//        //            //postSnmpIndex = row;
//        //            //string snmpcontent = testConfigRouterIntegrationTest[row, 1];
//        //            //SnmpReplaceContent[row] = snmpcontent;
//        //            //AppendLine2TextFile(cableConfigTxtFile, snmpcontent);
//        //            testConfigRouterIntegrationTest[row, 0] = (indexCurrentSnmp + 1).ToString();
//        //        }

//        //        /* Process the title */
//        //        if (testConfigRouterIntegrationTest[row, 0] == "")
//        //        {
//        //            continue;
//        //        }

//        //        if (testConfigRouterIntegrationTest[row, 0].ToLower().IndexOf("index") >= 0)
//        //        {
//        //            continue;
//        //        }                

//        //        if (!Int32.TryParse(testConfigRouterIntegrationTest[row, 0].Trim(), out indexCurrentSnmp))
//        //        {
//        //            bRouterIntegrationTestFTThreadRunning = false;
//        //            MessageBox.Show("Index Parsing Failed", "Error");
//        //            this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //            threadRouterIntegrationTestFT.Abort();
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

//        //            st_CbtMibDataConfigReadWriteTest[row].Name = testConfigRouterIntegrationTest[row, 1];
//        //            st_CbtMibDataConfigReadWriteTest[row].FullName = testConfigRouterIntegrationTest[row, 2];
//        //            st_CbtMibDataConfigReadWriteTest[row].Oid = testConfigRouterIntegrationTest[row, 3];
//        //            if(testConfigRouterIntegrationTest[row, 4] != "")
//        //                st_CbtMibDataConfigReadWriteTest[row].OidPlus = testConfigRouterIntegrationTest[row, 4];
//        //            st_CbtMibDataConfigReadWriteTest[row].Type = testConfigRouterIntegrationTest[row, 5];
//        //            st_CbtMibDataConfigReadWriteTest[row].AccessType = testConfigRouterIntegrationTest[row, 6];
//        //            st_CbtMibDataConfigReadWriteTest[row].Indexes = testConfigRouterIntegrationTest[row, 7];
//        //            //st_CbtMibDataConfigReadWriteTest[row].DefaultValue = testConfigRouterIntegrationTest[row, 8];
//        //            st_CbtMibDataConfigReadWriteTest[row].Module = testConfigRouterIntegrationTest[row, 8];
//        //            st_CbtMibDataConfigReadWriteTest[row].Description = testConfigRouterIntegrationTest[row, 9];
//        //            st_CbtMibDataConfigReadWriteTest[row].WrtieValue = testConfigRouterIntegrationTest[row, 10];
//        //            st_CbtMibDataConfigReadWriteTest[row].ExpectedValue = testConfigRouterIntegrationTest[row, 11];
//        //            int iParseValue;
//        //            if(!Int32.TryParse(testConfigRouterIntegrationTest[row, 12],out iParseValue))
//        //            {
//        //                st_CbtMibDataConfigReadWriteTest[row].WaitTime = 1 ;
//        //            }
//        //            else
//        //            {
//        //                st_CbtMibDataConfigReadWriteTest[row].WaitTime = iParseValue;
//        //            }
                    
//        //            //st_CbtMibDataConfigReadWriteTest[row].ReadValue = testConfigRouterIntegrationTest[row, 1]; ;
//        //            //st_CbtMibDataConfigReadWriteTest[row].Name = testConfigRouterIntegrationTest[row, 1]; ;


//        //            //string sName = testConfigRouterIntegrationTest[row, 1];
//        //            //string sFullName = testConfigRouterIntegrationTest[row, 2];
//        //            //string sOid = testConfigRouterIntegrationTest[row, 3];
//        //            //string SType = testConfigRouterIntegrationTest[row, 4];
//        //            //string sAccessType = testConfigRouterIntegrationTest[row, 5];
//        //            //string sIndexces = testConfigRouterIntegrationTest[row, 6];
//        //            //string sDefaleValue = testConfigRouterIntegrationTest[row, 7];
//        //            //string sMidModule = testConfigRouterIntegrationTest[row, 8];
//        //            //string sDescription = testConfigRouterIntegrationTest[row, 9];                    
//        //            //string sWriteValue = testConfigRouterIntegrationTest[row, 10];
//        //            //string sExpectedValue = testConfigRouterIntegrationTest[row, 11];

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
//        //            if (row != testConfigRouterIntegrationTest.GetLength(0) - 1)
//        //            {
//        //                row--;
//        //            }

//        //            /* Convert txt to bin file */
//        //            if (File.Exists(cableConfigBinFile))
//        //            {
//        //                str = "cfg File Exist! Delete File: " + cableConfigBinFile;
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //                File.Delete(cableConfigBinFile);
//        //            }

//        //            b_CfgNotExist = false;

//        //            RemoveEmptyLinesInFile_Common(cableConfigCfgTxtFile);

//        //            CableConfigFileTxt2Cfg(Path.GetDirectoryName(txtRouterIntegrationFunctionTestCfgConfigConverterFileName.Text), cableConfigCfgTxtFile, cableConfigBinFile, ref str);
//        //            if (str != "")
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //            if (!File.Exists(cableConfigBinFile))
//        //            {
//        //                //bRouterIntegrationTestFTThreadRunning = false;
//        //                //MessageBox.Show("Convert txt to Cfg Failed", "Error");
//        //                //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //                //threadRouterIntegrationTestFT.Abort();

//        //                str = "cfg File Convert fail. File doesn't Exist!!!! Test Next Index======";
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                                               
//        //                b_CfgNotExist = true;

//        //                //RouterIntegrationTestReportData(preSnmpIndex, postSnmpIndex, indexPreviousSnmp, false, str);

//        //                /* Reset CableConfigFile to Cable*/
//        //                File.Copy(cableconfigCfgFileTemplate, cableConfigCfgTxtFile, true);

//        //                preSnmpIndex = postSnmpIndex + 1;
//        //                indexPreviousSnmp++;

//        //                m_PositionRouterIntegrationTest++;

//        //                continue;                        
//        //            }

//        //            /* upload mib.cfg file to \\NFS address\file */
//        //            string NfsPath = txtRouterIntegrationFunctionTestMibNfsLocation.Text;

//        //            str = "Check if tftp server exists?";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //            if (Directory.Exists(NfsPath))
//        //            {
//        //                str = "Server exist. Copy bin file to tftp server: " + cableConfigBinFile;
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //                CopyFileToNfsFoler(cableConfigBinFile, NfsPath, true);
//        //                str = "Copy File Succeed.";
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //            }
//        //            else
//        //            {
//        //                str = "Server not found. Test Abort.";
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //                bRouterIntegrationTestFTThreadRunning = false;
//        //                MessageBox.Show("Server is not founded.", "Error");
//        //                this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //                threadRouterIntegrationTestFT.Abort();
//        //            }

//        //            /* Reset Device to get new config file for mib read/write */

//        //            //if (!ChangeCableDirecotory_Common("CM_Console>", comPortRouterIntegrationTest))
//        //            //{
//        //            //    bRouterIntegrationTestFTThreadRunning = false;
//        //            //    MessageBox.Show("Chang Cable Directory Failed", "Error");
//        //            //    this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //            //    threadRouterIntegrationTestFT.Abort();
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
//        //            //CableChanProcess_Common(nudRouterIntegrationFunctionTestGotoChannel.Value.ToString(), comPortRouterIntegrationTest);
//        //            str = "Change Cable Directory to CM/DOCSIS_CTL.";
//        //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //            CableChanDirectoryRouterIntegrationTest(nudRouterIntegrationFunctionTestGotoChannel.Value.ToString(), comPortRouterIntegrationTest);
                    
//        //            comPortRouterIntegrationTest.WriteLine("goto_ds " + nudRouterIntegrationFunctionTestGotoChannel.Value.ToString());

//        //            str = "";
//        //            sCheckConfigFile = "";
//        //            bReChanProcess = false;
//        //            b_TlvError = false;
//        //            stopWatch.Start();

//        //            /* Wait for device chan */
//        //            while (true)
//        //            {
//        //                if (bRouterIntegrationTestFTThreadRunning == false)
//        //                {
//        //                    bRouterIntegrationTestFTThreadRunning = false;
//        //                    MessageBox.Show("Abort test", "Error");
//        //                    this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //                    threadRouterIntegrationTestFT.Abort();
//        //                    // Never go here ...
//        //                }

//        //                if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestGotoWaitTime.Value) *1000)
//        //                {
//        //                    break;
//        //                }

                        
//        //                str = comPortRouterIntegrationTest.ReadLine();
//        //                if(str != "")
//        //                {
//        //                    strCompare += str;
//        //                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
//        //                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //                        CopyFileToNfsFoler(cableConfigBinFile, NfsPath, true);
//        //                        str = "Copy File Succeed.";
//        //                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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

//        //            //    bRouterIntegrationTestFTThreadRunning = false;
//        //            //    MessageBox.Show("Config File Name is wrong!! Check the server First!!\r\n" + sErrorMsg, "Error");
//        //            //    this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //            //    threadRouterIntegrationTestFT.Abort();
//        //            //}

//        //            /* Check if the device process re-chan action */
//        //            if (!bReChanProcess)
//        //            {
//        //                b_ReChanOK = false;
//        //                //bRouterIntegrationTestFTThreadRunning = false;
//        //                //MessageBox.Show("Re-Chand Process desn't execute!!", "Error");
//        //                //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //                //threadRouterIntegrationTestFT.Abort();
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
//        //                //RouterIntegrationTestReportData(preSnmpIndex, postSnmpIndex, indexPreviousSnmp, bCrash, str);
//        //            }
//        //            else
//        //            { // Check show cfg
//        //                str = "Change Cable Directory to CM/CM_Console/CM>.";
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                        
//        //                str = "";
//        //                ChangeToCableConsoleCM(ref str, comPortRouterIntegrationTest);

//        //                str = "Show cfg...";
//        //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //                str = "";
//        //                RunCableShowCfgRouterIntegrationTest(ref str, comPortRouterIntegrationTest);

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

//        //                    bRouterIntegrationTestFTThreadRunning = false;
//        //                    MessageBox.Show("Config File Name is wrong!! Check the server First!!\r\n" + sErrorMsg, "Error");
//        //                    this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //                    threadRouterIntegrationTestFT.Abort();
//        //                }

//        //                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //                //RouterIntegrationTestReportData(preSnmpIndex, postSnmpIndex, indexPreviousSnmp, bCrash, str);
//        //            }
                                        

//        //            /* Reset Device to get new config file for mib read/write */
//        //            //str = "Reset Dut:";
//        //            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //            //str = "Read Mib Type:";
//        //            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //            //string mibType = "";
//        //            //valueResponsed = "";
//        //            //bool setStatus = false;
//        //            //Snmp_RouterIntegrationTest.GetMibSingleValueByVersion(txtRouterIntegrationFunctionTestMibResetOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//        //            //if (mibType != "")
//        //            //{
//        //            //    str = "Mib Type: " + mibType;
//        //            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //            //    str = "Set Reset OID.";
//        //            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //            //    setStatus = Snmp_RouterIntegrationTest.SetMibSingleValueByVersion(txtRouterIntegrationFunctionTestMibResetOid.Text.Trim(), mibType, txtRouterIntegrationFunctionTestMibResetValue.Text, snmpVersion);
//        //            //}
//        //            //else
//        //            //{
//        //            //    str = "Set Reset OID.";
//        //            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //            //    setStatus = Snmp_RouterIntegrationTest.SetMibSingleValueByVersion(txtRouterIntegrationFunctionTestMibResetOid.Text.Trim(), dudRouterIntegrationFunctionTestMibResetMibType.Text, txtRouterIntegrationFunctionTestMibResetValue.Text, snmpVersion);
//        //            //}
//        //            ////str = "Set Finished, Now Get Value.";

//        //            ///* Wait Device to reboot */
//        //            //Thread.Sleep(3000);
//        //            //bool pingStatus = false;
//        //            //str = String.Format("Wait For Device to Reset for :" + nudRouterIntegrationFunctionTestPingTimeout.Value.ToString() + " Seconds");
//        //            //Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //            ////Stopwatch stopWatch = new Stopwatch();
//        //            //stopWatch.Reset();
//        //            //stopWatch.Stop();
//        //            //stopWatch.Start();

//        //            //while (true)
//        //            //{
//        //            //    if (bRouterIntegrationTestFTThreadRunning == false)
//        //            //    {
//        //            //        bRouterIntegrationTestFTThreadRunning = false;
//        //            //        MessageBox.Show("Abort test", "Error");
//        //            //        this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //            //        threadRouterIntegrationTestFT.Abort();
//        //            //        // Never go here ...
//        //            //    }

//        //            //    if (stopWatch.ElapsedMilliseconds > (Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//        //            //    {
//        //            //        break;
//        //            //    }

//        //            //    if (PingClient(mtbRouterIntegrationFunctionTestDeviceAddress.Text, 1000))
//        //            //    {
//        //            //        pingStatus = true;
//        //            //        break;
//        //            //    }

//        //            //    Thread.Sleep(1000);
//        //            //    str = String.Format(".");
//        //            //    Invoke(new SetTextCallBackT(SetTextC), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//        //            //}

//        //            //if (!pingStatus)
//        //            //{
//        //            //    bRouterIntegrationTestFTThreadRunning = false;
//        //            //    str = String.Format("\tPing Failed!!");
//        //            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //            //    MessageBox.Show("Ping Failed, Abort Test!!", "Error");
//        //            //    this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//        //            //    threadRouterIntegrationTestFT.Abort();
//        //            //}

//        //            //str = String.Format("\tPing Succeed!!");
//        //            //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//        //            //Run Main Function
//        //            //RouterIntegrationTestReportData(preSnmpIndex, postSnmpIndex, indexPreviousSnmp, bCrash);

//        //            /* Reset CableConfigFile to Cable*/
//        //            File.Copy(cableconfigCfgFileTemplate, cableConfigCfgTxtFile, true);
                    
//        //            preSnmpIndex = postSnmpIndex + 1;
//        //            indexPreviousSnmp++;
//        //        }
                
//        //        m_PositionRouterIntegrationTest++;
//        //    }

//        //    return true;
//        //}

//        /* The function is recorded report data */
//        private void RouterIntegrationTestReportData_Sanity(string[] sReport)
//        {
//            string str = string.Empty;
//            string valueResponsed = string.Empty;
//            string expectedValue = string.Empty;
//            string mibType = string.Empty;
//            int iIndexConsoleLog = 0;
//            string sConsoleLogFile = string.Empty;
//            //s_ErrorMsgForReport = string.Empty;
//            int iConditionTimeout = Convert.ToInt32(nudRouterIntegrationFunctionTestConditionTimeout.Value) * 1000;

//            bool bStatus = true;

//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();

//            try
//            {
//                //for (int reportIndex = 0; reportIndex < sReport.Length; reportIndex++)
//                //{                    
//                //    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, reportIndex + 2] = sReport[reportIndex];
//                //}

//                xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5] = sReport;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 6] = s_RouterIntegrationFinalReportInfo;
//                xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7] = sReport[5];


//                /* Change Color to PASS and Fail */
//                if (sReport == "PASS")
//                {
//                    xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
//                }
//                else if (sReport == "FILE")
//                {
//                    /* Write Console Log File as HyperLink in Excel report */
//                    xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7];
//                    xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sReport[5], Type.Missing, "ConsoleLog", "ConsoleLog");

//                    if (sComment == "") return;
//                    xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);

//                    int iBlank = -1;

//                    for (int i = 8; i < 20; i++)
//                    {
//                        string s = ((Excel.Range)xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, i]).Text;

//                        if (s == null || s == "")
//                        {
//                            iBlank = i;
//                            break;
//                        }                                            
//                    }

//                    if (iBlank > 0)
//                    {
//                        string FileName = Path.GetFileNameWithoutExtension(sComment);
//                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, iBlank] = sComment;
//                        xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, iBlank];
//                        //xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sComment, Type.Missing, "File", "File");
//                        xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sComment, Type.Missing, FileName, "FILE");
//                    }
//                }
//                else if (sReport == "PASSFILE")
//                {
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5] = "PASS";
//                    xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

//                    int iBlank = -1;

//                    for (int i = 8; i < 20; i++)
//                    {
//                        string s = ((Excel.Range)xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, i]).Text;
                        
//                        if (s == null || s == "")                            
//                        {
//                            iBlank = i;
//                            break;
//                        }                        
//                    }

//                    if (iBlank > 0)
//                    {
//                        string FileName = Path.GetFileNameWithoutExtension(sComment);
//                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, iBlank] = sComment;
//                        xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, iBlank];
//                        //xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sComment, Type.Missing, "File", "File");
//                        xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sComment, Type.Missing, FileName, "FILE");
//                    }               
                    
//                    //xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 8] = sComment;
//                    //xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 8];
//                    //xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sComment, Type.Missing, "File", "File");
//                }
//                else if (sReport == "FAILFILE")
//                {
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5] = "FAIL";
//                    xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

//                    string strTemp = sComment;
//                    string[] s = strTemp.Split(new string[] { "FILE::" }, StringSplitOptions.RemoveEmptyEntries);               
                    
//                    if (s.Length >= 2)
//                    {
//                        xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 6] = s[0];

//                        int iBlank = -1;

//                        for (int i = 8; i < 20; i++)
//                        {
//                            string ss = ((Excel.Range)xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, i]).Text;

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
//                            xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, iBlank] = s[1];
//                            xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, iBlank];
//                            //xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sComment, Type.Missing, "File", "File");
//                            xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, s[1], Type.Missing, FileName, "FILE");
//                        }      
                        
//                        //xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 8] = sComment;
//                        //xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 8];
//                        //xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sComment, Type.Missing, "File", "File");
//                    }                    
//                }
//                else
//                {
//                    xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
//                }

//                /* Write Console Log File as HyperLink in Excel report */
//                xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7];
//                xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sReport[5], Type.Missing, "ConsoleLog", "ConsoleLog");
//            }
//            catch (Exception ex )
//            {
//                //Debug.WriteLine(" Write Data to Excel File: " + ex.ToString());
//                str = "Write Data to Excel File Exception: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            }

//            try
//            {
//                xls_excelWorkBookRouterIntegrationTest.Save();
//            }
//            catch (Exception ex)
//            {
//                str = "Save Excel Error: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            }
//        }
        
//        /* For reference , This function want't be used */
//        private void RouterIntegrationTestReportData(string[] sReport)
//        {
//            string str = string.Empty;
//            string valueResponsed = string.Empty;
//            string expectedValue = string.Empty;
//            string mibType = string.Empty;
//            int iIndexConsoleLog = 0;
//            string sConsoleLogFile = string.Empty;
//            //s_ErrorMsgForReport = string.Empty;
//            int iConditionTimeout = Convert.ToInt32(nudRouterIntegrationFunctionTestConditionTimeout.Value) * 1000;

//            bool bStatus = true;

//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Start();

//            try
//            {
//                for (int reportIndex = 0; reportIndex < sReport.Length; reportIndex++)
//                {                    
//                    xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, reportIndex + 2] = sReport[reportIndex];
//                }

//                /* Change Color to PASS and Fail */
//                if (sReport == "PASS")
//                {
//                    xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
//                }
//                else
//                {
//                    xls_excelWorkSheetRouterIntegrationTest.Range[xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5], xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 5]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
//                }

//                /* Write Console Log File as HyperLink in Excel report */
//                xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 7];
//                xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, sReport[5], Type.Missing, "ConsoleLog", "ConsoleLog");
//            }
//            catch (Exception ex )
//            {
//                //Debug.WriteLine(" Write Data to Excel File: " + ex.ToString());
//                str = "Write Data to Excel File Exception: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            }

//            try
//            {
//                xls_excelWorkBookRouterIntegrationTest.Save();
//            }
//            catch (Exception ex)
//            {
//                str = "Save Excel Error: " + ex.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            }

//        }

//        private string createRouterIntegrationTestSubFolder(string ModelName)
//        {
//            string subFolder = ((ModelName == "") ? "CGA2121_" : ModelName + "_") + DateTime.Now.ToString("yyyyMMdd-HHmmss");

//            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
//                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);

//            subFolder = System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder;
//            return subFolder;
//        }

//        private string createRouterIntegrationTestSavePath(string subFolder, ModelInfo info, int Loop)
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

//        private string createRouterIntegrationTestSaveExcelFile(string subFolder, ModelInfo info, int Loop, string FileName)
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

//        private string createRouterIntegrationTestSaveExcelFile(string subFolder, string Sku, int Loop, string FileName)
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

//        private string createRouterIntegrationTestVeriwaveSaveExcelFile(string subFolder, int Loop, string FileName)
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

//        //private bool CheckNeededParameter_RouterIntegrationTest()
//        //{
//        //    /* Config and Check Test Condition */
//        //    SetText("Check Test Condition...", txtRouterIntegrationFunctionTestInformation);
            
//        //    /* Check Test Condition Content */
//        //    if (dgvRouterIntegrationExcelTestConditionData.RowCount <= 1)
//        //    {
//        //        MessageBox.Show("Test Condition can't be Empty!!");
//        //        return false;
//        //    }

//        //    /* Check SNMP Parameter */
//        //    if (txtRouterIntegrationFunctionTestCfgFileName.Text == "")
//        //    {
//        //        MessageBox.Show("Cfg File Name can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (txtRouterIntegrationFunctionTestBinFileName.Text == "")
//        //    {
//        //        MessageBox.Show("Bin File Name can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (txtRouterIntegrationFunctionTestCfgHeaderFile.Text == "")
//        //    {
//        //        MessageBox.Show("Cfg Heafer File can't be Empty!!");
//        //        return false;
//        //    }

//        //    if(!File.Exists(txtRouterIntegrationFunctionTestCfgHeaderFile.Text))
//        //    {
//        //        MessageBox.Show("Cfg Header File doesn't Exist!!");
//        //        return false;
//        //    }

//        //    if (txtRouterIntegrationFunctionTestBinHeaderFile.Text == "")
//        //    {
//        //        MessageBox.Show("Bin Heafer File can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (!File.Exists(txtRouterIntegrationFunctionTestBinHeaderFile.Text))
//        //    {
//        //        MessageBox.Show("Bin Header File doesn't Exist!!");
//        //        return false;
//        //    }
            
//        //    if (txtRouterIntegrationFunctionTestCfgConfigConverterFileName.Text == "")
//        //    {
//        //        MessageBox.Show("Cfg Config converter File can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (!File.Exists(txtRouterIntegrationFunctionTestCfgConfigConverterFileName.Text))
//        //    {
//        //        MessageBox.Show("Cfg Config converter File doesn't Exist!!");
//        //        return false;
//        //    }

//        //    if (txtRouterIntegrationFunctionTestBinConfigConverterFileName.Text == "")
//        //    {
//        //        MessageBox.Show("Bin Config converter File can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (!File.Exists(txtRouterIntegrationFunctionTestBinConfigConverterFileName.Text))
//        //    {
//        //        MessageBox.Show("Bin Config converter File doesn't Exist!!");
//        //        return false;
//        //    }

//        //    if (txtRouterIntegrationFunctionTestMibNfsLocation.Text == "")
//        //    {
//        //        MessageBox.Show("NFS Location can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (mtbRouterIntegrationFunctionTestCmIpAddress.Text == "")
//        //    {
//        //        MessageBox.Show("CM IP Address can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (mtbRouterIntegrationFunctionTestMtaIpAddress.Text == "")
//        //    {
//        //        MessageBox.Show("MTA IP Address can't be Empty!!");
//        //        return false;
//        //    }

//        //    if (!CheckIPValid(mtbRouterIntegrationFunctionTestCmIpAddress.Text))
//        //    {
//        //        MessageBox.Show("CM IP Address is Invalid!!");
//        //        return false;
//        //    }

//        //    if (!CheckIPValid(mtbRouterIntegrationFunctionTestMtaIpAddress.Text))
//        //    {
//        //        MessageBox.Show("MTA IP Address is Invalid!!");
//        //        return false;
//        //    }

//        //    //QuickPingRouterIntegrationTest(mtbRouterIntegrationFunctionTestCmIpAddress.Text);

//        //    //QuickPingRouterIntegrationTest(mtbRouterIntegrationFunctionTestMtaIpAddress.Text);

//        //    testConfigRouterIntegrationTest = null;

//        //    /* config file pre-process - convert header file for backup ,especiallly device crash */
//        //    /* CM */
//        //    string str = string.Empty;

//        //    /* Convert default txt to basic cfg file */
//        //    if (File.Exists(s_CableConfigCfgBasicCfgFile))
//        //    {
//        //        SetText("Delete CFG Basic Config File.", txtRouterIntegrationFunctionTestInformation);
//        //        File.Delete(s_CableConfigCfgBasicCfgFile);
//        //    }

//        //    CableConfigFileTxt2Cfg(Path.GetDirectoryName(txtRouterIntegrationFunctionTestCfgConfigConverterFileName.Text), txtRouterIntegrationFunctionTestCfgHeaderFile.Text, s_CableConfigCfgBasicCfgFile, ref str);
//        //    SetText(str, txtRouterIntegrationFunctionTestInformation);

//        //    if (!File.Exists(s_CableConfigCfgBasicCfgFile))
//        //    {
//        //        MessageBox.Show("Convert Basic Config File Failed.");
//        //        return false;
//        //    }

//        //    /* Copy uesr header txt file to txt template */
//        //    if (File.Exists(cableconfigCfgFileTemplate))
//        //    {
//        //        SetText("Delete Header Template File.", txtRouterIntegrationFunctionTestInformation);
//        //        File.Delete(cableconfigCfgFileTemplate);
//        //    }

//        //    File.Copy(txtRouterIntegrationFunctionTestCfgHeaderFile.Text, cableconfigCfgFileTemplate, true);     

//        //    /* Copy txt temp file to txt file */
//        //    File.Copy(cableconfigCfgFileTemplate, cableConfigCfgTxtFile, true);   

//        //    /* MTA */
//        //    //TODO , 
            





//        //     SetText("Check result: OK.", txtRouterIntegrationFunctionTestInformation);
//        //    return true;

//        //    //if (rbtnRouterIntegrationFunctionTestMibExcelFile.Checked)
//        //    //{
//        //    //    // Check Excel File Exist, and Read Excel File
//        //    //    if (txtRouterIntegrationFunctionTestMibExcelFile.Text == "")
//        //    //    {
//        //    //        MessageBox.Show("Excel file can't be Empty!!");
//        //    //        return false;
//        //    //    }

//        //    //    if (txtRouterIntegrationFunctionTestCfgHeaderFile.Text == "")
//        //    //    {
//        //    //        MessageBox.Show("CFG Header file can't be Empty!!");
//        //    //        return false;
//        //    //    }

//        //    //    if (!File.Exists(txtRouterIntegrationFunctionTestMibExcelFile.Text))
//        //    //    {
//        //    //        MessageBox.Show("Excel file doesn't exist!!");
//        //    //        return false;

//        //    //    }

//        //    //    if (!File.Exists(txtRouterIntegrationFunctionTestCfgHeaderFile.Text))
//        //    //    {
//        //    //        MessageBox.Show("CFG Header file doesn't exist!!");
//        //    //        return false;
//        //    //    }

//        //    //    testConfigRouterIntegrationTest = null;

//        //    //    //if (!File.Exists(txtRouterIntegrationFunctionTestMibExcelFile.Text))
//        //    //    //{
//        //    //    //    MessageBox.Show("Excel file doesn't exist!!");
//        //    //    //    return false;
//        //    //    //}

//        //    //    //if (!ReadTestConfig_ReadExcelRouterIntegrationTest(txtRouterIntegrationFunctionTestMibExcelFile.Text))
//        //    //    //{
//        //    //    //    MessageBox.Show("Read Excel file Failed!!");
//        //    //    //    return false;
//        //    //    //}
//        //    //}           
//        //}

//        private bool QuickPingRouterIntegrationTest(string ip, int iTimeout = 3000)
//        {
//            /* Ping if the deivce available*/
//            bool pingStatus = false;
//            //SetText("Ping Device...", txtRouterIntegrationFunctionTestInformation);
//            Stopwatch stopWatch = new Stopwatch();
//            stopWatch.Reset();
//            stopWatch.Stop();
//            stopWatch.Start();

//            while (true)
//            {
//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    bRouterIntegrationTestFTThreadRunning = false;
//                    //MessageBox.Show("Abort test", "Error");
//                    //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                    //threadRouterIntegrationTestFT.Abort();
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
//                //SetText("Ping Failed!!", txtRouterIntegrationFunctionTestInformation);
//                //MessageBox.Show("Ping Failed, Abort Test!!", "Error");
//                return false;
//            }

//            //SetText("Ping Succeed!!", txtRouterIntegrationFunctionTestInformation);
//            return true;
//        }
        
//        //private void InitialParameter_RouterIntegrationTest()
//        //{
//        //    testConfigRouterIntegrationTest = null;
//        //    xlApp = null;
//        //    xlWorkbook = null;
//        //    xlWorksheet = null;
//        //    xlRange = null;
//        //    st_CbtMibDataConfigReadWriteTest = null;

//        //    txtRouterIntegrationFunctionTestInformation.Text = "";
//        //    return;
//        //}

//        private void ToggleRouterIntegrationFunctionTestGUI()
//        {
//            ToggleRouterIntegrationFunctionTestController(true);
//            Debug.WriteLine("Toggle");
//        }

//        private bool ReadTestConfig_ReadExcelRouterIntegrationTest(string ExcelFile)
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

//            testConfigRouterIntegrationTest = rowdata;

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

//            //    testConfigRouterIntegrationTest = newRowData;
//            //}
//            //else
//            //{
//            //    testConfigRouterIntegrationTest = rowdata;
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

//        private bool ReadSanityTestConfig_ReadExcelRouterIntegrationTest(string ExcelFile)
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

//            testConfigRouterIntegrationSanityTest = rowdata;

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

//            //    testConfigRouterIntegrationSanityTest = newRowData;
//            //}
//            //else
//            //{
//            //    testConfigRouterIntegrationSanityTest = rowdata;
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

//        private bool CableChanDirectoryRouterIntegrationTest(string sChannel, Comport cComport)
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
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            while (true)
//            {
//                if (bRouterIntegrationTestFTThreadRunning == false)
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
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                Thread.Sleep(1000);
//            }

//            return true;
//        }




//        private bool RunCableShowCfgRouterIntegrationTest(ref string sInfo, Comport cComport)
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
//            //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            //    sInfo += str;
//            //}

//            while (true)
//            {
//                if (bRouterIntegrationTestFTThreadRunning == false)
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
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    sInfo += str;
//                }

//                str = cComport.ReadLine();                

//                //Thread.Sleep(100);
//            }

//            return true;
//        }




//        private bool CGA2121RebootDut_RouterIntegrationTest()
//        {
//            string str = string.Empty;
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds
//            Stopwatch stopWatch = new Stopwatch();

//            //CGA2121BackToRootDirectory_RouterIntegrationTest("cm");

//            str = "reset Device...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            comPortRouterIntegrationTest.WriteLine("/reset");
//            stopWatch.Start();

//            while (true)
//            {
//                if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestDutRebootTimeout.Value * 1000))
//                {
//                    //str = "Reboot Timeout.";
//                    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    //return false;
//                    break;
//                }

//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    return false;
//                }

//                str = comPortRouterIntegrationTest.ReadLine();

//                if (str != "")
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //if (str.IndexOf(sCableDownloadCfg) != -1)
//                //{
//                //    str = "Reboot Finished.";
//                //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                //    return true;
//                //}
//            }

//            return true;
//        }





//        private bool CGA2121ShowCmCfg_RouterIntegrationTest()
//        {
//            string str = string.Empty;
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds
//            Stopwatch stopWatch = new Stopwatch();            

//            // Check show cfg
//            str = "Check show cfg...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            if (!CGA2121BackToRootDirectory_RouterIntegrationTest("cm"))
//            {
//                return false;
//            }

//            bool bCfgStatus = false;
//            s_CfgContect = "";
//            str = "";
//            comPortRouterIntegrationTest.DiscardBuffer();
//            Thread.Sleep(1000);
//            comPortRouterIntegrationTest.WriteLine("/Console/cm/show cfg");

//            str = comPortRouterIntegrationTest.ReadLine();
//            if (str != "")
//            {
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                s_CfgContect += str;
//            }

//            stopWatch.Stop();
//            stopWatch.Reset();
//            stopWatch.Restart();

//            while (true)
//            {
//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    bRouterIntegrationTestFTThreadRunning = false;
//                    return false;
//                    // Never go here ...
//                }

//                if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//                {
//                    str = "Show cfg Timeout.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    s_CfgContect += str;
//                }

//                str = comPortRouterIntegrationTest.ReadLine();
//                //Thread.Sleep(100);
//            }

//            stopWatch.Stop();
//            stopWatch.Reset();

//            str = "Check Config File Name...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            /* Check if the device loaded cfg file file is equal to CableConfigFile.cfg */
//            string sFwFileName = RouterIntegrationDutDataSets[i_RouterIntegrationCurrentDateSetsIndex].CurrentFwName;

//            str = "FW File Name should be: " + sFwFileName;
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                return false;
//            }

//            return true;            
//        }

//        private bool CGA2121CheckFwFileName_RouterIntegrationTest(int iIndex)
//        {
//            string str = string.Empty;
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds
//            Stopwatch stopWatch = new Stopwatch();

//            // Check show cfg
//            str = "Show Firmware File Name...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //if (!CGA2121BackToRootDirectory_RouterIntegrationTest("cm"))
//            //{
//            //    return false;
//            //}

//            bool bCfgStatus = false;
//            s_CfgContect = "";
//            str = "";
//            comPortRouterIntegrationTest.DiscardBuffer();
//            Thread.Sleep(1000);

//            comPortRouterIntegrationTest.WriteLine("/ver");

//            str = comPortRouterIntegrationTest.ReadLine();
//            if (str != "")
//            {
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                s_CfgContect += str;
//            }

//            stopWatch.Stop();
//            stopWatch.Reset();
//            stopWatch.Restart();

//            while (true)
//            {
//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    bRouterIntegrationTestFTThreadRunning = false;
//                    return false;
//                    // Never go here ...
//                }

//                if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//                {
//                    str = "Show version Timeout.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
//                    comPortRouterIntegrationTest.WriteLine("");
//                    //break;
//                }

//                if (str != "" && bCfgStatus)
//                {
//                    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    s_CfgContect += str;
//                }

//                str = comPortRouterIntegrationTest.ReadLine();
//                if (str != "")
//                {
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                }
//            }

//            stopWatch.Stop();
//            stopWatch.Reset();

//            str = "Check Firmware File Name...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            /* Check if the device loaded cfg file file is equal to CableConfigFile.cfg */
//            string sFwFileName = RouterIntegrationDutDataSets[iIndex].CurrentFwName;

//            str = "FW File Name should be: " + sFwFileName;
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                return false;
//            }

//            return true;
//        }

//        private bool CGA2121RebootBySnmpReadByComport_RouterIntegrationTest()
//        {
//            string str = string.Empty;
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds
//            Stopwatch stopWatch = new Stopwatch();

//            CGA2121BackToRootDirectory_RouterIntegrationTest("cm");

//            str = "reset Device...";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            comPortRouterIntegrationTest.WriteLine("/reset");
//            stopWatch.Start();

//            while (true)
//            {
//                if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestDutRebootTimeout.Value * 1000))
//                {
//                    str = "Reboot Timeout.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    return false;
//                }

//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    return false;
//                }

//                str = comPortRouterIntegrationTest.ReadLine();

//                if (str != "")
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //if (str.IndexOf(sCableDownloadCfg) != -1)
//                //{
//                //    str = "Reboot Finished.";
//                //    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                //    return true;
//                //}
//            }

//            return true;
//        }
        
//        private bool RestoreBasicConfigFileAndRebootCableByConsoleRouterIntegrationTest()
//        {
//            string str = string.Empty;
//            Stopwatch stopWatch = new Stopwatch();

//            //Copy Basic Config File to tftp Server
//            //if(s_TestType.ToLower() == "cm")
//            //{
//            //    RestoreCfgBasicConfigFileRouterIntegrationTest();
//            //    //Reboot DUT

//            //}

            
//            ChangeToCableConsoleCM(ref str, comPortRouterIntegrationTest);
//            comPortRouterIntegrationTest.WriteLine("/reset");
//            stopWatch.Start();            

//            while (true)
//            {
//                if(stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestDutRebootTimeout.Value *1000))
//                {
//                    str = "Reboot Timeout.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    return false;
//                }

//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    return false;
//                }

//                str = comPortRouterIntegrationTest.ReadLine();

//                if(str != "") 
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                if (str.IndexOf(sCableDownloadCfg) != -1)
//                {
//                    str = "Reboot Finished.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    return true;                        
//                }
//            }
//        }

//        private bool RestoreCfgBasicConfigFileRouterIntegrationTest()
//        {
//            string str = string.Empty;
//            string NfsPath = "";// txtRouterIntegrationFunctionTestMibNfsLocation.Text;
//            string sBasicCfgFile = "";//txtRouterIntegrationFunctionTestCfgFileName.Text;

//            //File.Copy(s_CableConfigCfgBasicCfgFile, cableConfigBinFile, true);
//            File.Copy(s_CableConfigCfgBasicCfgFile, sBasicCfgFile, true);
            
//            str = " Copy basic cfg file to tftp server.";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            CopyFileToNfsFoler(sBasicCfgFile, NfsPath, true);
//            str = "Copy File Succeed.";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            return true;
//        }

        private bool ParameterCheckAndInitial_RouterIntegrationTest()
        {
            /* Initialize  */
//            testConfigRouterIntegrationTest = null;
//            xlApp = null;
//            xlWorkbook = null;
//            xlWorksheet = null;
//            xlRange = null;
//            st_CbtMibDataConfigReadWriteTest = null;

//            txtRouterIntegrationFunctionTestInformation.Text = "";
//            bRouterIntegrationThroughputTest = false;
//            t_Timer = null;
//            Snmp_RouterIntegrationTest = null;

//            //i_SetCheckHourTime = 5;
//            //i_SetCheckMinuteTime = 00;
//            //i_SetCheckSecondTime = 00;

//            /* Check the day now is less than the end day */
//            if (DateTime.Compare(dtpRouterIntegrationConfigurationFtpTestPeriodStartDay.Value, dtpRouterIntegrationConfigurationFtpTestPeriodEndDay.Value) > 0)
//            {
//                MessageBox.Show("Start Day can't be late than End Day!!");
//                return false;
//            }

//            if (DateTime.Compare(DateTime.Now, dtpRouterIntegrationConfigurationFtpTestPeriodEndDay.Value) > 0)
//            {
//                MessageBox.Show("Today is not in the Test Period!!");
//                return false;
//            }

//            /* Config and Check Test Condition */
//            SetText("Check Paramters...", txtRouterIntegrationFunctionTestInformation);

//            /* Check Test Condition Content */
//            if (dgvRouterIntegrationDutsSettingData.RowCount <= 1)
//            {
//                MessageBox.Show("Device Setting Content can't be Empty!!");
//                return false;
//            }
            
//            /* Check SNMP Parameter */
//            if (mtbRouterIntegrationConfigurationCmtsCasaC10gIpAddress.Text == "" || mtbRouterIntegrationConfigurationCmtsArrisC4IpAddress.Text == "")
//            {
//                MessageBox.Show("CMTS server IP Address can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationRdFtpServerPath.Text == "" || txtRouterIntegrationConfigurationNfsServerPath.Text == "")
//            {
//                MessageBox.Show("Ftp server Path can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationReportEmailSendTo.Text == "")
//            {
//                MessageBox.Show("Email Send to list can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationTestConditionExcelFile.Text == "")
//            {
//                MessageBox.Show("Test Condition Excel File Name can't be Empty!!");
//                return false;
//            }            

//            if (!CheckIPValid(mtbRouterIntegrationConfigurationCmtsCasaC10gIpAddress.Text) || !CheckIPValid(mtbRouterIntegrationConfigurationCmtsArrisC4IpAddress.Text))
//            {
//                MessageBox.Show("CMTS Server IP Address is Invalid!!");
//                return false;
//            }

//            //Device SNMP Content Check 
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Description OID can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationResetOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Reset OID can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationResetValue.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Reset Value can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationTftpServerOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Tftp Server OID can't be Empty!!");
//                return false;
//            }

//            if (mtbRouterIntegrationConfigurationServerIp.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Tftp Server IP Value can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationImageFileOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: FW Image File OID can't be Empty!!");
//                return false;
//            }

//            if (dudRouterIntegrationConfigurationResetType.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Reset Type can't be Empty!!");
//                return false;
//            }
            
//            if (dudRouterIntegrationConfigurationAdminStatusType.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Admin Status Type can't be Empty!!");
//                return false;
//            }

//            if (dudRouterIntegrationConfigurationTftpServerTypeType.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Tftp Server Type Type can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationAdminStatusOID.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Admin Status OID can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationAdminStatusValue.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Admin Status Value can't be Empty!!");
//                return false;
//            }
            
//            if (txtRouterIntegrationConfigurationTftpServerTypeOID.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Tftp Server Type OID can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationTftpServerTypeValue.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Tftp Server Type Value can't be Empty!!");
//                return false;
//            }

//            if (nudRouterIntegrationConfigurationInProcessTimeout.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: In-Process Timeout Value can't be Empty!!");
//                return false;
//            }
            
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }
//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationDescriptionOid.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content Description Oid can't be Empty!!");
//                return false;
//            }

//            if (nudRouterIntegrationConfigurationFwUpdateWaitTime.Text == "")
//            {
//                MessageBox.Show("Device SNMP Content: Wait FW Download Time Value can't be Empty!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationCfgConverterFileName.Text == "")
//            {
//                MessageBox.Show("Configuration Cfg Converter File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtRouterIntegrationConfigurationCfgConverterFileName.Text))
//            {
//                MessageBox.Show("Configuration Cfg Converter File doesn't exist!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationCfgSnmpEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("Snmp Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtRouterIntegrationConfigurationCfgSnmpEuTxtFileName.Text))
//            {
//                MessageBox.Show("Snmp Euro Txt File doesn't exist!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationCfgTlvEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("Tlv Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtRouterIntegrationConfigurationCfgTlvEuTxtFileName.Text))
//            {
//                MessageBox.Show("Tlv Euro Txt File doesn't exist!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationCfgP7bEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("P7b Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtRouterIntegrationConfigurationCfgP7bEuTxtFileName.Text))
//            {
//                MessageBox.Show("P7b Euro Txt doesn't exist!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationCfgSnmpNonEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("Snmp Non-Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtRouterIntegrationConfigurationCfgSnmpNonEuTxtFileName.Text))
//            {
//                MessageBox.Show("Snmp Non-Euro Txt File doesn't exist!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationCfgTlvNonEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("Tlv Non-Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtRouterIntegrationConfigurationCfgTlvNonEuTxtFileName.Text))
//            {
//                MessageBox.Show("Tlv Non-Euro Txt File doesn't exist!!");
//                return false;
//            }

//            if (txtRouterIntegrationConfigurationCfgP7bNonEuTxtFileName.Text == "")
//            {
//                MessageBox.Show("P7b Non-Euro Txt File can't be Empty!!");
//                return false;
//            }
//            if (!File.Exists(txtRouterIntegrationConfigurationCfgP7bNonEuTxtFileName.Text))
//            {
//                MessageBox.Show("P7b Non-Euro Txt File doesn't exist!!");
//                return false;
//            }
            
//            SetText("Done!!", txtRouterIntegrationFunctionTestInformation);
//            SetText("Read DUT Settings...", txtRouterIntegrationFunctionTestInformation);

//            /* Read Device Data Setting */
//            ReadDataSets_RouterIntegrationTest();

//            if (bRouterIntegrationThroughputTest)
//            {
//                //if (btnRouterIntegrationVeriwaveTestConditionEditSetting.Text.ToLower() == "cancel")
//                //{
//                //    btnRouterIntegrationVeriwaveTestConditionEditSetting.Text = "Edit";
//                //    hasDeleteButton = false;
//                //    dgvRouterIntegrationVeriwaveTestConditionData.Columns.Remove("Action");
//                //}

//                //if (dgvRouterIntegrationVeriwaveTestConditionData.RowCount > 1)
//                //{
//                //    //if (MessageBox.Show("The Data in the list will be deleted", "Warning", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
//                //    //    return;
//                //    //else
//                //    //{
//                //    //    DataTable dt = (DataTable)dgvRouterIntegrationVeriwaveTestConditionData.DataSource;
//                //        dgvRouterIntegrationVeriwaveTestConditionData.Rows.Clear();
//                //        //dgvRouterIntegrationVeriwaveTestConditionData.DataSource = dt;
//                //    //}
//                //}
//                //readXmlRouterIntegrationVeriwaveTestCondition(System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterIntegrationVeriwaveTestCondition.xml");

//                if (txtRouterIntegrationConfigurationWaveappFolder.Text == "")
//                {
//                    MessageBox.Show("Veriwave Wave app Folder can't be Empty!!");
//                    return false;
//                }

//                if (!File.Exists(txtRouterIntegrationConfigurationWaveappFolder.Text))
//                {
//                    MessageBox.Show("Veriwave Wave app File doesn't Exist!!");
//                    return false;
//                }

//                if (txtRouterIntegrationConfigurationReportFolder.Text == "")
//                {
//                    MessageBox.Show("Veriwave Rrport Folder can't be Empty!!");
//                    return false;
//                }

//                if (!Directory.Exists(txtRouterIntegrationConfigurationReportFolder.Text))
//                {
//                    try
//                    {
//                        string[] saDirTemp = txtRouterIntegrationConfigurationReportFolder.Text.Split('\\');
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

//                if (!Directory.Exists(txtRouterIntegrationConfigurationReportFolder.Text))
//                {
//                    MessageBox.Show("Veriwave Report Folder doesn't exist!!");
//                    return false;
//                }

//                if (dgvRouterIntegrationVeriwaveTestConditionData.RowCount <= 1)
//                {
//                    MessageBox.Show("Veriwave Test Condition can't be Empty!!");
//                    return false;
//                }

//                /* Set Veriwave app cl foleder */
//                string waveFile = txtRouterIntegrationConfigurationWaveappFolder.Text;
//                string wavePath = waveFile.Substring(0, waveFile.Length - Path.GetFileName(waveFile).Length);
//                s_WaveappclFile = wavePath + "waveapps_cl.exe";

//                /* Check waveapps_cl.exe */
//                //SetText("Check waveapps_cl.exe...", txtRouterIntegrationFunctionTestInformation);
//                if (!File.Exists(s_WaveappclFile))
//                {
//                    MessageBox.Show("waveapps_cl.exe doesn't exist!!");
//                    return false;
//                }

//                sa_VeriwaveOrigReportFile = null;
//                s_VeriwaveReportFolder = txtRouterIntegrationConfigurationReportFolder.Text;
//                string[] sa_temp;
//                GetDirectoryFolders_CommonFunction(txtRouterIntegrationConfigurationReportFolder.Text, out sa_temp);
//                sa_VeriwaveOrigReportFile = new string[sa_temp.Length + dgvRouterIntegrationVeriwaveTestConditionData.RowCount];
//                for (int i = 0; i < sa_temp.Length; i++)
//                {
//                    sa_VeriwaveOrigReportFile[i] = sa_temp[i];
//                }

//                //SetText("Check Succeed!", txtRouterIntegrationFunctionTestInformation);
//            }         
            
//            /* Read Test Condition */
//            string sExcelFileName = txtRouterIntegrationConfigurationTestConditionExcelFile.Text;
//            string str = string.Empty;

//            if (!File.Exists(txtRouterIntegrationConfigurationTestConditionExcelFile.Text))
//            {
//                MessageBox.Show("Test Condition Excel File doesn't Exist!!");
//                return false;
//            }

//            if (!ReadSanityTestConfig_ReadExcelRouterIntegrationTest(sExcelFileName))
//            {
//                MessageBox.Show("Read Excel Failed!!");
//                return false;
//            }            

//            sExcelFileName = System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterIntegrationTestSanityPreVersion.xlsx";

//            if (!File.Exists(sExcelFileName))
//            {
//                MessageBox.Show("Test Condition Excel File doesn't Exist!!");
//                return false;
//            }

//            if (!ReadTestConfig_ReadExcelRouterIntegrationTest(sExcelFileName))
//            {
//                MessageBox.Show("Read Excel Failed!!");
//                return false;
//            }

//            s_EmailSenderID = txtRouterIntegrationConfigurationReportEmailSenderGmailAccount.Text;
//            s_EmailSenderPassword = txtRouterIntegrationConfigurationReportEmailSenderGmailPassword.Text;
//            s_EmailReceivers = txtRouterIntegrationConfigurationReportEmailSendTo.Text.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            
//            SetText("Done!!", txtRouterIntegrationFunctionTestInformation);

//            if (!ReadDeviceSsidSecurityModePresharedKey_RouterIntegrationTest())
//            {
//                //MessageBox.Show("Read SSID Security Key or Preshared Key Failed!!");
//                return false;
//            }

            return true;
        }

//        private bool ReadDataSets_RouterIntegrationTest()
//        {
//            RouterIntegrationDutDataSets[] sDataSets = new RouterIntegrationDutDataSets[dgvRouterIntegrationDutsSettingData.RowCount - 1];
//            //string[,] dataSets = new string[dgvRouterIntegrationDutsSettingData.RowCount - 1, dgvRouterIntegrationDutsSettingData.ColumnCount];

//            for (int i = 0; i < dgvRouterIntegrationDutsSettingData.RowCount - 1; i++)
//            {
//                sDataSets[i].index = Convert.ToInt32(dgvRouterIntegrationDutsSettingData.Rows[i].Cells[0].Value.ToString());
//                sDataSets[i].IpAddress = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[1].Value.ToString();
//                sDataSets[i].MacAddress = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[2].Value.ToString();
//                sDataSets[i].PcIpAddress = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[3].Value.ToString();
//                sDataSets[i].SwitchPort = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[4].Value.ToString();
//                sDataSets[i].ComportNum = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[5].Value.ToString();
//                sDataSets[i].TestType = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[6].Value.ToString();
//                sDataSets[i].CmtsType = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[7].Value.ToString();
//                sDataSets[i].SkuModel = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[8].Value.ToString();
//                sDataSets[i].SerialNumber = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[9].Value.ToString();
//                sDataSets[i].SwVersion = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[10].Value.ToString();
//                sDataSets[i].HwVersion = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[11].Value.ToString();
//                sDataSets[i].FwFileName = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[12].Value.ToString();
//                sDataSets[i].SSID = dgvRouterIntegrationDutsSettingData.Rows[i].Cells[13].Value.ToString();
//            }

//            RouterIntegrationDutDataSets = sDataSets;

//            int snmpVersion = rbtnRouterIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnRouterIntegrationConfigurationSnmpV2.Checked ? 2 : 3;

//            try
//            {
//                for (int i = 0; i < RouterIntegrationDutDataSets.Length; i++)
//                {
//                    RouterIntegrationDutDataSets[i].snmp = new CBTSnmp();
//                    RouterIntegrationDutDataSets[i].snmp.Init(RouterIntegrationDutDataSets[i].IpAddress, snmpVersion, dudRouterIntegrationConfigurationSnmpReadCommunity.Text, dudRouterIntegrationConfigurationSnmpWriteCommunity.Text, Convert.ToInt32(nudRouterIntegrationConfigurationSnmpPort.Value), Convert.ToInt32(nudRouterIntegrationConfigurationTrapPort.Value));
//                }
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show("Snmp Initial Failed!!: " + ex.ToString());
//                System.Windows.Forms.Cursor.Current = Cursors.Default;
//                btnRouterIntegrationFunctionTestRun.Enabled = true;
//                return false;
//            }

//            foreach (RouterIntegrationDutDataSets set in sDataSets)
//            {
//                if (set.TestType.ToLower().IndexOf("throughput") >= 0)
//                {
//                    bRouterIntegrationThroughputTest = true;                    
//                }
//            }            

//            return true;
//        }

//        private bool ReadDeviceSsidSecurityModePresharedKey_RouterIntegrationTest()
//        {
//            string str = string.Empty;
//            SetText("Read SSID, Security Mode, Wpa PreShared key before Test!!", txtRouterIntegrationFunctionTestInformation);

//            /* Read TchRgDot11BssSsid , TchRgDot11BssSecurityMode , TchRgDot11WpaPreSharedkey Before test */
//            /* OID: TchRgDot11BssSsid , TchRgDot11BssSecurityMode , TchRgDot11WpaPreSharedkey */
//            //const string cs_CGA2121BssSsid24gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.2.1.1.3.32";
//            //const string cs_CGA2121BssSsid5gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.2.1.1.3.112";
//            //const string cs_CGA2121BssSecurityMode24gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.2.1.1.4.32";
//            //const string cs_CGA2121BssSecurityMode5gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.2.1.1.4.112";
//            //const string cs_CGA2121WpaPreSharedkey24gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.3.1.1.2.32";
//            //const string cs_CGA2121WpaPreSharedkey5gOid = ".1.3.6.1.4.1.46366.4292.79.2.2.3.1.1.2.112";

//            for (int i = 0; i < RouterIntegrationDutDataSets.Length; i++)
//            {
//                //if (bRouterIntegrationTestFTThreadRunning == false)
//                //{
//                //    //MessageBox.Show("Abort!!");
//                //    return false;
//                //}

//                string sIp = RouterIntegrationDutDataSets[i].IpAddress;
//                str = "Ping Device First: " + sIp;
//                SetText(str, txtRouterIntegrationFunctionTestInformation);

//                ////Ping Device First
//                ///* Ping ip first */
//                ////if (!QuickPingRouterIntegrationTest(mtbRouterIntegrationFunctionTestCmIpAddress.Text, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//                //if (!QuickPingRouterIntegrationTest(sIp, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//                //{
//                //    str = String.Format("Sku: {0} Ping Device Failed.", RouterIntegrationDutDataSets[i].SkuModel);
//                //    SetText(str, txtRouterIntegrationFunctionTestInformation);

//                //    MessageBox.Show("Ping Device Failed!!");
//                //    return false;
//                //}

//                /* Ping if the deivce available*/
//                bool pingStatus = false;
//                int iTimeout = 5 * 1000; //5 Seconds
//                //SetText("Ping Device...", txtRouterIntegrationFunctionTestInformation);
//                Stopwatch stopWatch = new Stopwatch();
//                stopWatch.Reset();
//                stopWatch.Stop();
//                stopWatch.Start();

//                while (true)
//                {
//                    //if (bRouterIntegrationTestFTThreadRunning == false)
//                    //{
//                    //    bRouterIntegrationTestFTThreadRunning = false;
//                    //    //MessageBox.Show("Abort test", "Error");
//                    //    //this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                    //    //threadRouterIntegrationTestFT.Abort();
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
//                    SetText("Ping Failed!!", txtRouterIntegrationFunctionTestInformation);
//                    str = String.Format("Sku: {0}, IP:{1} Ping Failed!!", RouterIntegrationDutDataSets[i].SkuModel, sIp);
//                    MessageBox.Show(str, "Error");
//                    return false;
//                }

//                SetText("Ping Succeed!!", txtRouterIntegrationFunctionTestInformation);

//                //if (bRouterIntegrationTestFTThreadRunning == false)
//                //{
//                //    //MessageBox.Show("Abort!!");
//                //    return false;
//                //}

//                string valueResponsed = string.Empty;
//                string mibType = string.Empty;
//                int snmpVersion = rbtnRouterIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnRouterIntegrationConfigurationSnmpV2.Checked ? 2 : 3;

//                try
//                {
//                    /* 2.4G SSID */
//                    str = String.Format("Sku: {0} 2.4G SSID.", RouterIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterIntegrationFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121BssSsid24gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} SSID is not Valid.", RouterIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterIntegrationFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    RouterIntegrationDutDataSets[i].PreSsid24G = valueResponsed.Trim();

//                    //if (bRouterIntegrationTestFTThreadRunning == false)
//                    //{
//                    //    //MessageBox.Show("Abort!!");
//                    //    return false;
//                    //}

//                    /* 5G SSID */
//                    str = String.Format("Sku: {0} 5G SSID.", RouterIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterIntegrationFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121BssSsid5gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} SSID is not Valid.", RouterIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterIntegrationFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    RouterIntegrationDutDataSets[i].PreSsid5G = valueResponsed.Trim();

//                    //if (bRouterIntegrationTestFTThreadRunning == false)
//                    //{
//                    //    //MessageBox.Show("Abort!!");
//                    //    return false;
//                    //}

//                    /* 2.4G Security Mode */
//                    str = String.Format("Sku: {0} 2.4G Security Mode.", RouterIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterIntegrationFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121BssSecurityMode24gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} Security Mode is not Valid.", RouterIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterIntegrationFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    RouterIntegrationDutDataSets[i].PreSecurityMode24G = valueResponsed.Trim();

//                    //if (bRouterIntegrationTestFTThreadRunning == false)
//                    //{
//                    //    //MessageBox.Show("Abort!!");
//                    //    return false;
//                    //}

//                    /* 5G Security Mode */
//                    str = String.Format("Sku: {0} 5G Security Mode.", RouterIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterIntegrationFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121BssSecurityMode5gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} Security Mode is not Valid.", RouterIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterIntegrationFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    RouterIntegrationDutDataSets[i].PreSecurityMode5G = valueResponsed.Trim();

//                    //if (bRouterIntegrationTestFTThreadRunning == false)
//                    //{
//                    //    //MessageBox.Show("Abort!!");
//                    //    return false;
//                    //}

//                    /* 2.4G PreShared Key */
//                    str = String.Format("Sku: {0} 2.4G PreShared Key.", RouterIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterIntegrationFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121WpaPreSharedkey24gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} PreShared Key is not Valid.", RouterIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterIntegrationFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    RouterIntegrationDutDataSets[i].PreWpaPreSharedkey24G = valueResponsed.Trim();

//                    //if (bRouterIntegrationTestFTThreadRunning == false)
//                    //{
//                    //    //MessageBox.Show("Abort!!");
//                    //    return false;
//                    //}

//                    /* 5G PreShared Key */
//                    str = String.Format("Sku: {0} 5G PreShared Key.", RouterIntegrationDutDataSets[i].SkuModel);
//                    SetText(str, txtRouterIntegrationFunctionTestInformation);

//                    valueResponsed = "";
//                    mibType = "";

//                    RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(cs_CGA2121WpaPreSharedkey5gOid, ref valueResponsed, ref mibType, snmpVersion);

//                    str = "Mib Type: " + mibType;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = String.Format("Read Value: {0}", valueResponsed);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    if (valueResponsed == null || valueResponsed == "")
//                    {
//                        str = String.Format("Sku: {0} PreShared Key is not Valid.", RouterIntegrationDutDataSets[i].SkuModel);
//                        SetText(str, txtRouterIntegrationFunctionTestInformation);
//                        MessageBox.Show(str);
//                        return false;
//                    }

//                    RouterIntegrationDutDataSets[i].PreWpaPreSharedkey5G = valueResponsed.Trim();
//                }
//                catch (Exception ex)
//                {
//                    str = String.Format("Sku: {0} Read SSId, Security Mode or PreShared Key Exception: {1}", RouterIntegrationDutDataSets[i].SkuModel, ex.ToString());
//                    SetText(str, txtRouterIntegrationFunctionTestInformation);
//                    MessageBox.Show(str);
//                    return false;
//                }
//            }

//            SetText("Done!!", txtRouterIntegrationFunctionTestInformation);

//            return true;
//        }
        
//        private bool CGA2121BackToRootDirectory_RouterIntegrationTest(string mode)
//        {
//            string str = string.Empty;
//            int iChangeDirectoryTimeout = 60 * 1000; // 60 seconds

//            /* Create a time check */
//            Stopwatch stopWatch = new Stopwatch();

//            comPortRouterIntegrationTest.DiscardBuffer();
//            Thread.Sleep(1000);

//            comPortRouterIntegrationTest.WriteLine(" cd /");
//            comPortRouterIntegrationTest.WriteLine("");
//            str = comPortRouterIntegrationTest.ReadLine();
//            if (str != "")
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
            
//            stopWatch.Stop();
//            stopWatch.Reset();
//            stopWatch.Restart();

//            while (true)
//            {
//                if (bRouterIntegrationTestFTThreadRunning == false)
//                    return false;

//                if (stopWatch.ElapsedMilliseconds > iChangeDirectoryTimeout)
//                {
//                    str = "Change Directory Timeout.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    return false;
//                }

//                if (mode.ToLower() == "cm")
//                {
//                    if (str.ToLower().IndexOf("cm>") >=0 ) break;

//                    if (str.ToLower().IndexOf("rg>") <0)
//                    {
//                        comPortRouterIntegrationTest.WriteLine("/Console/sw");
//                        str = comPortRouterIntegrationTest.ReadLine();
//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        Thread.Sleep(3000);
//                        comPortRouterIntegrationTest.WriteLine("");
//                        str = comPortRouterIntegrationTest.ReadLine();
//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        comPortRouterIntegrationTest.WriteLine("cd /");
//                    }
//                }

//                if (mode.ToLower() == "mta")
//                {
//                    if (str.ToLower().IndexOf("rg>") >=0) break;

//                    if (str.ToLower().IndexOf("cm>") <0)
//                    {
//                        comPortRouterIntegrationTest.WriteLine("/Console/sw");
//                        str = comPortRouterIntegrationTest.ReadLine();
//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        Thread.Sleep(3000);
//                        comPortRouterIntegrationTest.WriteLine("");
//                        str = comPortRouterIntegrationTest.ReadLine();
//                        if (str != "")
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        comPortRouterIntegrationTest.WriteLine("cd /");
//                    }
//                }

//                comPortRouterIntegrationTest.WriteLine("");
//                str = comPortRouterIntegrationTest.ReadLine();
//            }          

//            return true;
//        }

//        private bool CreateSkuFoldRouterIntegrationTest()
//        {
//             bool isExists;
//             string subPath = m_RouterIntegrationTestSubFolder;

//            isExists = System.IO.Directory.Exists(subPath);
//            if (!isExists)
//                System.IO.Directory.CreateDirectory(subPath);

//            //if (bRouterIntegrationThroughputTest)
//            //{
//            //    subPath = subPath + "\\Veriwave";
//            //    RouterIntegrationVeriwave.ReportSavePath = subPath;
//            //    isExists = System.IO.Directory.Exists(subPath);
//            //    if (!isExists)
//            //        System.IO.Directory.CreateDirectory(subPath);
//            //}

//            for (int i = 0; i < RouterIntegrationDutDataSets.Length; i++)
//            {
//                subPath = subPath + "\\" + RouterIntegrationDutDataSets[i].SkuModel;
//                RouterIntegrationDutDataSets[i].ReportSavePath = subPath;
//                isExists = System.IO.Directory.Exists(subPath);
//                if (!isExists)
//                    System.IO.Directory.CreateDirectory(subPath);
//            }
            
//            return true;
//        }

//        private bool CreateSubFoldRouterIntegrationVeriwave()
//        {
//            bool isExists;
//            string subPath = m_RouterIntegrationTestSubFolder;

//            isExists = System.IO.Directory.Exists(subPath);
//            if (!isExists)
//                System.IO.Directory.CreateDirectory(subPath);

//            subPath = subPath + "\\Veriwave";
//            RouterIntegrationVeriwave.ReportSavePath = subPath;
//            isExists = System.IO.Directory.Exists(subPath);
//            if (!isExists)
//                System.IO.Directory.CreateDirectory(subPath);
           
//            return true;
//        }
        
//        private bool CopyRdFwToNfsServerRouterIntegrationTest()
//        {
//            string str = string.Empty;
//            //string[] sSku = new string[] { "CHILE", "COLOMBIA", "PERU", "GERMANY" };
//            List<string> list = new List<string>();
//            bool bRdFwDownloadStatus = true;
//            string[] sInTempFile = new string[30];

//            string sDay = DateTime.Now.ToString("yyyy-MM-dd");
//            string sShorDay = DateTime.Now.ToString("yyMMdd");

//            //string RdPath = "192.168.65.201";
//            string RdPath = txtRouterIntegrationConfigurationRdFtpServerPath.Text;
//            string RdTempPath = System.Windows.Forms.Application.StartupPath + "\\Temp";
//            string sDirOrig = "/Technicolor/Taipei_BFC5.7.1mp3_RG/CGA2121";
//            string sDir = sDirOrig;
//            //string sFwFileNameOid = ".1.3.6.1.2.1.69.1.3.2.0";
//            //string sFwFileNameOidType = "OctetString";
//            //string[] sFwFiles = new string[RouterIntegrationDutDataSets.Length];
            

//            str = "Date is " + sDay;
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            for (int i = 0; i < RouterIntegrationDutDataSets.Length; i++)
//            {
//                sDir = sDirOrig;
//                //sDir = sDir + "/CGA2121_" + sSku[i] + "/";
//                if (RouterIntegrationDutDataSets[i].SkuModel.IndexOf("GERMANY") >= 0)
//                {
//                    sDir = sDir + "/CGA2121_GERMANY/";
//                }
//                else
//                {
//                    sDir = sDir + "/CGA2121_" + RouterIntegrationDutDataSets[i].SkuModel + "/";
//                }
//                sDir = RdPath + sDir;

//                str = "Ftp Path is: ftp://" + sDir;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        bRdFwDownloadStatus = false;                        

//                        //超過早上10點都沒有FW,就隔天再做.
//                        if (i_SetCheckHourTime >= 10)
//                        {
//                            i_SetCheckHourTime = 5;

//                            str = "Try again Tomorrow morning!!!";
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                            i_SetCheckHourTime++;

//                            str = "Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                            return false;
//                        }

//                        str = "Wait for one hour and try again!!!";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        i_SetCheckHourTime++;

//                        str = "Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        return false;
//                    }                   

//                    str = "The Date Folder Exists: ftp://" + sDir;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    ftpclient = null;
//                    ftpclient = new CbtFtpClient(sDir, "cbtsqa", "jenkins");

//                    /* Get File list in today's folder */
//                    list.Clear();
//                    list = ftpclient.GetFtpFileList();
//                    string[] sFileName = new string[20];
//                    //iIndex = 0;
//                    string sDownloadFileName = RouterIntegrationDutDataSets[i].FwFileName;
//                    sDownloadFileName = sDownloadFileName.Replace("#DATE", sShorDay);
//                    RouterIntegrationDutDataSets[i].CurrentFwName = sDownloadFileName;

//                    int iIndex = 0;

//                    //列出目錄中所有檔案
//                    str = "List all the files in Folder: "  + sDir;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    foreach (string s in list)
//                    {
//                        string[] sFiles = s.Split(' ');
//                        foreach (string st in sFiles)
//                        {
//                            if (st.IndexOf("CGA") >= 0)
//                            {
//                                Invoke(new SetTextCallBackT(SetText), new object[] { st, txtRouterIntegrationFunctionTestInformation });
//                                sFileName[iIndex++] = st;
//                            }
//                        }
//                    }
                    
//                    //確認檔案是否存在
//                    str = "Check if the file exists :" + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        bRdFwDownloadStatus = false;                      

//                        str = "Wait for one hour and try again!!!";
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                        i_SetCheckHourTime++;

//                        str = "Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
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
//                    //                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    //            }
//                    //        }
//                    //    }
//                    //}

//                    //iIndex = 0;

//                    ftpclient = null;
//                    ftpclient = new CbtFtpClient(RdPath, "cbtsqa", "jenkins");
//                    str = "Start to Download File: " + sDownloadFileName;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = "";
//                    ftpclient.DownloadFile(RdTempPath, sDir, sDownloadFileName, ref str);
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
//                    //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    //        str = "";
//                    //        ftpclient.DownloadFile(RdTempPath, sDir, st, ref str);
//                    //        Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    //        str = "Download File Succeed: " + st;
//                    //        sInTempFile[iIndex++] = st;
//                    //    }
//                    //}
//                    ftpclient = null;
//                }
//                catch (Exception ex)
//                {
//                    str = "RD FW Doanload Failed: " + ex.ToString();
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    str = "Wait 1 Minutes and try Again !!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    Thread.Sleep(60 *1000);                   
//                    //str = "Start Over!!!";
//                    //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    i = -1;
//                }
//            }

//            // RD FW 下載完成  
//            str = "RD FW Doanload Finished!!! Start to Upload to NFS Server!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });            

//            string NfsPath = txtRouterIntegrationConfigurationNfsServerPath.Text;
//            //上傳 Firmware 檔案到NFS Server 上面去.

//            str = "Check if Local tftp server exists?";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            if (!Directory.Exists(NfsPath))
//            { //連不到NFS的資料夾,等一個小時之後再試
//                str = "Local Tftp Server is unreachable!!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                bRdFwDownloadStatus = false;
//                str = "Wait for one hour and try again!!!";
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                i_SetCheckHourTime++;

//                str = "Set Test Time to: " + i_SetCheckHourTime.ToString() + ":" + i_SetCheckMinuteTime.ToString() + ":" + i_SetCheckSecondTime.ToString();
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                return false;

//                //bRouterIntegrationTestFTThreadRunning = false;
//                //MessageBox.Show("Tftp Server is unreachable!!!", "Error");
//                ////this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));
//                //threadRouterIntegrationTestFT.Abort();
//                //if (thread_RouterIntegrationVeriwaveFT == null || thread_RouterIntegrationVeriwaveFT.ThreadState == System.Threading.ThreadState.Stopped)
//                //    this.Invoke(new showRouterIntegrationTestGUIDelegate(ToggleRouterIntegrationFunctionTestGUI));

                
//                //str = "The Date Folder doesn't exist!! ftp://" + sDir + sDay;
//                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            }

//            str = "Server exist. Copy file to local tftp server: ";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            foreach (string s in sInTempFile)
//            {
//                if (s != null)
//                {
//                    string sFwFile = RdTempPath + "\\" + s;
//                    //string NfsPath = "\\10.6.76.7\tftp root";
//                    //string NfsPath = "";
//                    str = "Copy File to Local NFS server: " + sFwFile;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    string sDestFile = NfsPath + "\\" + s;

//                    //CopyFileToNfsFoler(sFwFile, NfsPath, true);
//                    File.Copy(sFwFile, sDestFile, true);

//                    str = "Copy File Succeed.";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                    ////刪除已下載的FW
//                    //if (File.Exists(sFwFile))
//                    //{
//                    //    File.Delete(sFwFile);
//                    //}
//                }
//            }

//            str = "Finished!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            str = "Start to Write FW File Name to each Device Process!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            //將要Download的 FW 檔名寫入Device中
//            int snmpVersion = rbtnRouterIntegrationConfigurationSnmpV1.Checked ? 1 : rbtnRouterIntegrationConfigurationSnmpV2.Checked ? 2 : 3;

//            for (int i = 0; i < RouterIntegrationDutDataSets.Length; i++)
//            {
//                string valueResponsed = string.Empty;
//                string expectedValue = string.Empty;
//                string mibType = string.Empty;

//                string sIp = RouterIntegrationDutDataSets[i].IpAddress;

//                str = "Ping Device First: " + sIp;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                 //Ping Device First
//                /* Ping ip first */
//                //if (!QuickPingRouterIntegrationTest(mtbRouterIntegrationFunctionTestCmIpAddress.Text, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//                if (!QuickPingRouterIntegrationTest(sIp, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//                {
//                    str = " Failed!! ";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    //string saveFileLog = savePath.Substring(0, savePath.Length - 4) + "_Header.log";                    
//                    bRdFwDownloadStatus = false;
//                    str = "Wait for one hour and try again!!!";
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                    i_SetCheckHourTime++;
//                    continue;
//                    //return false;
//                }

//                //Setting Tftp Server IP
//                str = "Setting Tftp Server IP to Device: " + RouterIntegrationDutDataSets[i].SkuModel + ", " + RouterIntegrationDutDataSets[i].CurrentFwName;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });                
//                str = "Tftp Server IP: " + mtbRouterIntegrationConfigurationServerIp.Text;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerOid.Text.Trim(), mibType, mtbRouterIntegrationConfigurationServerIp.Text.Trim(), snmpVersion);
//                Thread.Sleep(1000);
//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //Setting Tftp Server Address
//                str = "Setting Tftp Server Address Type: " + txtRouterIntegrationConfigurationTftpServerTypeValue.Text;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerTypeOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerTypeOID.Text.Trim(), mibType, txtRouterIntegrationConfigurationTftpServerTypeValue.Text.Trim(), snmpVersion);
//                Thread.Sleep(1000);
//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationTftpServerTypeOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                
//                str = "Write Fw File Name to Device: " + RouterIntegrationDutDataSets[i].SkuModel + ", " + RouterIntegrationDutDataSets[i].CurrentFwName;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                valueResponsed = "";
//                mibType = "";
//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationImageFileOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationImageFileOid.Text.Trim(), mibType, RouterIntegrationDutDataSets[i].CurrentFwName, snmpVersion);
//                Thread.Sleep(1000);
                
//                // 讀回FW Name
//                valueResponsed = "";
//                mibType = "";

//                RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationImageFileOid.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
                
//                str = "Mib Type: " + mibType;
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
                
//                str = String.Format("Read Value: {0}", valueResponsed);
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(sFwFileNameOid, "OctetString", sFwFiles[i], snmpVersion);
//                //setStatus = Snmp_CableMibReadWriteTest.SetMibSingleValueByVersion(sOid, mibType, sWriteValue, snmpVersion);              
            
//                //// Set Admin Status to 1 for downloading firmware
//                //valueResponsed = "";
//                //mibType = "";
//                //RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);
//                //RouterIntegrationDutDataSets[i].snmp.SetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), mibType, txtRouterIntegrationConfigurationAdminStatusValue.Text.Trim(), snmpVersion);
//                //Thread.Sleep(1000);
//                //valueResponsed = "";
//                //mibType = "";
//                //RouterIntegrationDutDataSets[i].snmp.GetMibSingleValueByVersion(txtRouterIntegrationConfigurationAdminStatusOID.Text.Trim(), ref valueResponsed, ref mibType, snmpVersion);

//                //str = "Mib Type: " + mibType;
//                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//                //str = String.Format("Read Value: {0}", valueResponsed);
//                //Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });                
//            }

//            str = "Finished!!";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

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
//            //if (!QuickPingRouterIntegrationTest(mtbRouterIntegrationFunctionTestCmIpAddress.Text, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//            //if (!QuickPingRouterIntegrationTest(sIp, Convert.ToInt32(nudRouterIntegrationFunctionTestPingTimeout.Value * 1000)))
//            //{

                
//                //        
//                //        File.WriteAllText(saveFileLog, txtRouterIntegrationFunctionTestInformation.Text);
//                //        Invoke(new ShowTextboxContentDelegate(ShowTextboxContent), new object[] { "", txtRouterIntegrationFunctionTestInformation });

//                //        try
//                //        {                   
//                //            xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 2] = "Ping Failed";                  

//                //            /* Write Console Log File as HyperLink in Excel report */
//                //            xls_excelRangeRouterIntegrationTest = xls_excelWorkSheetRouterIntegrationTest.Cells[m_PositionRouterIntegrationTest, 3];
//                //            xls_excelWorkSheetRouterIntegrationTest.Hyperlinks.Add(xls_excelRangeRouterIntegrationTest, saveFileLog, Type.Missing, "ConsoleLog", "ConsoleLog");               
//                //        }
//                //        catch (Exception ex)
//                //        {
//                //            str = "Write Data to Excel File Exception: " + ex.ToString();
//                //            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                //        }
//                //        /* Save end time and close the Excel object */
//                //        xls_excelAppRouterIntegrationTest.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
//                //        xls_excelWorkBookRouterIntegrationTest.Save();
//                //        /* Close excel application when finish one job(power level) */
//                //        closeExcelRouterIntegrationTest();
//                //        Thread.Sleep(3000);               
//                //        return false;
//            //}
            
//            // 準備開始Device測試, 下次下載FW時間設定為明天早上5:00
//            i_SetCheckHourTime = 5;
//            i_SetCheckMinuteTime = 00;
//            i_SetCheckSecondTime = 00;
           
//            return true;
//        }
        
//        private bool DownloadFwAndRebootBySnmpRouterIntegrationTest()
//        {






//            return true;
//        }

//        //private bool RebootDeviceRouterIntegrationTest(string sChannel, Comport cComport)
//        //{
//        //    return true;
//        //}

//        private bool CGA2121CmReChanProcess_RouterIntegrationTest(int Freq)
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
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                return false;
//            }

//            str = "ReChan Device with Freq " + Freq.ToString();
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            /* use dogo_ds freq command to re-chan device */
//            //str = "Run goto_ds " + nudCableMibConfigReadWriteFunctionTestGotoChannel.Value.ToString();
//            str = "Run goto_ds 333";
//            Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });

//            str = "";

//            stopWatch.Stop();
//            stopWatch.Reset();
//            stopWatch.Restart();

//            long lTimeTemp = stopWatch.ElapsedMilliseconds + 30 * 1000;

//            comPortRouterIntegrationTest.WriteLine("/docsis_ctl/goto_ds " + Freq.ToString());
//            str = comPortRouterIntegrationTest.ReadLine();
//            if (str != "")
//            {
//                Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//            }

//            while (true)
//            {
//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    bRouterIntegrationTestFTThreadRunning = false;
//                    return false;
//                    // Never go here ...
//                }

//                if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestDutRebootTimeout.Value) * 1000)
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
//                    comPortRouterIntegrationTest.WriteLine("/docsis_ctl/goto_ds " + Freq.ToString());
//                }
//                str = comPortRouterIntegrationTest.ReadLine();
//                if (str != "")
//                {
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                }
//            }

//            if (!b_ReChanOK) return false;

//            stopWatch.Stop();
//            stopWatch.Reset();
//            stopWatch.Restart();

//            /* Wait for device chan */
//            while (true)
//            {
//                if (bRouterIntegrationTestFTThreadRunning == false)
//                {
//                    bRouterIntegrationTestFTThreadRunning = false;
//                    return false;
//                    // Never go here ...
//                }

//                if (stopWatch.ElapsedMilliseconds > Convert.ToInt32(nudRouterIntegrationFunctionTestDutRebootTimeout.Value) * 1000)
//                {
//                    break;
//                }

//                str = comPortRouterIntegrationTest.ReadLine();
//                if (str != "")
//                {
//                    strCompare += str;
//                    Invoke(new SetTextCallBackT(SetText), new object[] { str, txtRouterIntegrationFunctionTestInformation });
//                }

//                if (strCompare.ToLower().IndexOf("crash") != -1 || strCompare.ToLower().IndexOf("bcm338498") != -1)
//                {
//                    str = "========== System Crash ============";
//                    return false;
//                }
//            }

//            return true;
//        }

//        private bool CGA2121SanityTest_RouterIntegrationTest(int iStartIndex, int iStopIndex)
//        {
//            for (int i = iStartIndex; i <= iStopIndex; i++)
//            {
                

//            }
//            return true;
//        }

//        private void LoadRouterIntegrationDataGridView()
//        {
//            string sFileName = System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterIntegrationDutsSetting.xml";
//            if(File.Exists(sFileName))
//            {
//                readXmlRouterIntegrationDutsSetting(sFileName);
//                //readXmlRouterIntegrationDutsSetting(System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterIntegrationDutsSetting.xml");
//                //readXmlRouterIntegrationVeriwaveTestCondition(System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterIntegrationVeriwaveTestCondition.xml");
//            }

//            sFileName = System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterIntegrationVeriwaveTestCondition.xml";
//            if (File.Exists(sFileName))
//            {
//                readXmlRouterIntegrationVeriwaveTestCondition(sFileName);
//                //readXmlRouterIntegrationVeriwaveTestCondition(System.Windows.Forms.Application.StartupPath + "\\testCondition\\RouterIntegrationVeriwaveTestCondition.xml");
//            }

//        }


//        #endregion





        //private void WebGUIMainFunctionIntegration()
        //{
        //    Thread.Sleep(2000);
        //    CBT_SeleniumApi.GuiScriptParameter ScriptPara = new CBT_SeleniumApi.GuiScriptParameter();
        //    string sExceptionInfo = string.Empty;
        //    string sLoginURL = string.Format("http://{0}:{1}", ls_LoginSettingParametersIntegration.GatewayIP, ls_LoginSettingParametersIntegration.HTTP_Port);
        //    string sCurrentURL = string.Empty;
        //    int DATA_ROW = 15;
        //    bool bTestResult = true;
        //    int j = i_TestScriptIndexIntegration;

        //    ScriptPara.Procedure = st_ReadScriptDataIntegration[j].Procedure;
        //    ScriptPara.Index = st_ReadScriptDataIntegration[j].TestIndex;
        //    ScriptPara.Action = st_ReadScriptDataIntegration[j].Action;
        //    ScriptPara.ActionName = st_ReadScriptDataIntegration[j].ActionName;
        //    ScriptPara.ElementType = st_ReadScriptDataIntegration[j].ElementType;
        //    ScriptPara.ElementXpath = st_ReadScriptDataIntegration[j].ElementXpath;
        //    ScriptPara.ElementXpath = ScriptPara.ElementXpath.Replace('\"', '\'');
        //    ScriptPara.RadioBtnExpectedValueXpath = st_ReadScriptDataIntegration[j].RadioButtonExpectedValueXpath.Replace('\"', '\'');
        //    ScriptPara.WriteValue = st_ReadScriptDataIntegration[j].WriteExpectedValue;
        //    ScriptPara.ExpectedValue = st_ReadScriptDataIntegration[j].WriteExpectedValue;
        //    ScriptPara.URL = sLoginURL + st_ReadScriptDataIntegration[j].WriteExpectedValue;
        //    ScriptPara.TestTimeOut = st_ReadScriptDataIntegration[j].TestTimeOut;
        //    ScriptPara.Note = string.Empty;

        //    ---------------------------------------//
        //    -------------- Go To URL --------------//
        //    ---------------------------------------//
        //    if (j == 0)
        //    {
        //        s_CurrentURLIntegration = sLoginURL;
        //        cs_BrowserIntegration.GoToURL(sLoginURL);
        //        Thread.Sleep(1000);
        //    }
        //    if (ScriptPara.Action.CompareTo("Goto") == 0 && s_CurrentURLIntegration.CompareTo(ScriptPara.URL) != 0)
        //    {
        //        #region Go To URL
        //        s_CurrentURLIntegration = ScriptPara.URL;
        //        Thread.Sleep(1000);
        //        cs_BrowserIntegration.GoToURL(s_CurrentURLIntegration);
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    -------------- Set Value --------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("Set") == 0)
        //    {
        //        #region Set Value
        //        s_InfoStrIntegration = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Set Value", ScriptPara.Index, ScriptPara.ActionName);
        //        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrIntegration, txtIntegrationFunctionTestInformation });

        //        sw_TestTimerIntegration.Stop();
        //        bool checkXPathResult = false;
        //        Stopwatch waitingXpath = new Stopwatch();
        //        waitingXpath.Start();
        //        do
        //        {
        //            checkXPathResult = cs_BrowserIntegration.CheckXPathDisplayed(ScriptPara.ElementXpath);
        //        } while (waitingXpath.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);

        //        if (checkXPathResult == true)
        //        {
        //            Thread.Sleep(2000);
        //        }
        //        else if (checkXPathResult == false)
        //        {
        //            s_InfoStrIntegration = string.Format("...Couldn't find the element!");
        //            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrIntegration, txtIntegrationFunctionTestInformation });
        //        }
        //        waitingXpath.Reset();
        //        sw_TestTimerIntegration.Start();
        //        try
        //        {
        //            if (ScriptPara.WriteValue != "" && st_ReadStepsScriptDataIntegration[i_StepsScriptIndexIntegration].WriteValue[i_CommandDataIndex] != "" && st_ReadStepsScriptDataIntegration[i_StepsScriptIndexIntegration].WriteValue[i_CommandDataIndex] != null)
        //            {
        //                ScriptPara.WriteValue = st_ReadStepsScriptDataIntegration[i_StepsScriptIndexIntegration].WriteValue[i_CommandDataIndex];
        //                i_CommandDataIndex++;
        //            }
        //        }
        //        catch { }
        //        ScriptPara.Note = string.Empty;
        //        try
        //        {
        //            cs_BrowserIntegration.SetWebElementValue(ref ScriptPara);
        //        }
        //        catch
        //        {
        //            ExceptionActionIntegration(s_CurrentURLIntegration, ref ScriptPara);
        //            ScriptPara.Note = "Set Value Error:\n" + ScriptPara.Note;
        //            SwitchToRGconsoleIntegration();
        //            WriteconsoleLogIntegration();
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
        //        s_InfoStrIntegration = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Get Value", ScriptPara.Index, ScriptPara.ActionName);
        //        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrIntegration, txtIntegrationFunctionTestInformation });

        //        ScriptPara.Note = string.Empty;
        //        ScriptPara.GetValue = string.Empty;

        //        if (ScriptPara.ElementType.CompareTo("RADIO_BUTTON") == 0)
        //        {
        //            ScriptPara.ElementXpath = st_ReadScriptDataIntegration[j].RadioButtonExpectedValueXpath.Replace('\"', '\'');
        //        }

        //        try
        //        {
        //            bTestResult = cs_BrowserIntegration.GetWebElementValue(ref ScriptPara); // Set Value
        //        }
        //        catch
        //        {
        //            ExceptionActionIntegration(s_CurrentURLIntegration, ref ScriptPara);
        //            ScriptPara.Note = "Execute SubmitButton Error:\n" + ScriptPara.Note;
        //            return false;
        //        }


        //        ---------- Write Test Report ----------//
        //        WriteTestReportIntegration(j, TestResult, ref DATA_ROW, ScriptPara.Note);
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    --------------- ReLogin ---------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("ReLogin") == 0)
        //    {
        //        #region ReLogin
        //        Invoke(new SetTextCallBack(SetText), new object[] { "Log in again...", txtIntegrationFunctionTestInformation });
        //        Re_LoginIntegration();
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
        //        Invoke(new SetTextCallBack(SetText), new object[] { "Login...", txtIntegrationFunctionTestInformation });
        //        while (true)
        //        {
        //            if (pingTimer.ElapsedMilliseconds > (Convert.ToInt32(120000)))
        //            {
        //                break;
        //            }

        //            if (PingClient(ls_LoginSettingParametersIntegration.GatewayIP, 1000))
        //            {
        //                Invoke(new SetTextCallBack(SetText), new object[] { "\tPing Successfully!!", txtIntegrationFunctionTestInformation });
        //                break;
        //            }

        //            Thread.Sleep(1000);
        //            string Info = String.Format(".");
        //            Invoke(new SetTextCallBack(SetText), new object[] { Info, txtIntegrationFunctionTestInformation });
        //        }
        //        Thread.Sleep(3000);
        //        String[] splitLoginInfo = ScriptPara.WriteValue.Split('/');
        //        cs_BrowserIntegration.loginAlertMessage(ScriptPara.URL, splitLoginInfo[0], splitLoginInfo[1]);
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
        //        s_InfoStrIntegration = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:File Upload", ScriptPara.Index, ScriptPara.ActionName);
        //        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrIntegration, txtIntegrationFunctionTestInformation });

        //        Thread.Sleep(3000);
        //        ScriptPara.Note = string.Empty;
        //        try
        //        {
        //            cs_BrowserIntegration.fileUploadE8350(ref ScriptPara);
        //        }
        //        catch
        //        {
        //            ExceptionActionIntegration(s_CurrentURLIntegration, ref ScriptPara);
        //            ScriptPara.Note = "Set Value Error:\n" + ScriptPara.Note;
        //            SwitchToRGconsoleIntegration();
        //            WriteconsoleLogIntegration();
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
        //            s_InfoStrIntegration = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Waiting for {2} seconds", ScriptPara.Index, ScriptPara.ActionName, ScriptPara.TestTimeOut);
        //            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrIntegration, txtIntegrationFunctionTestInformation });
        //            sw_TestTimerIntegration.Stop();
        //            Thread.Sleep(Convert.ToInt16(ScriptPara.TestTimeOut) * 1000);
        //            sw_TestTimerIntegration.Start();
        //            SwitchToRGconsoleIntegration();
        //            WriteconsoleLogIntegration();
        //        }
        //        else
        //        {
        //            s_InfoStrIntegration = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Waiting for the XPath, it won't be more than {2} seconds", ScriptPara.Index, ScriptPara.ActionName, ScriptPara.TestTimeOut);
        //            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrIntegration, txtIntegrationFunctionTestInformation });
        //            sw_TestTimerIntegration.Stop();
        //            bool checkXPathResult = false;
        //            Stopwatch waitingXpath = new Stopwatch();
        //            waitingXpath.Start();
        //            do
        //            {
        //                checkXPathResult = cs_BrowserIntegration.CheckXPathDisplayed(ScriptPara.ElementXpath);
        //            } while (waitingXpath.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);
        //            waitingXpath.Reset();
        //            sw_TestTimerIntegration.Start();
        //            SwitchToRGconsoleIntegration();
        //            WriteconsoleLogIntegration();
        //        }
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    --------------- CloseDriver ---------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("CloseDriver") == 0)
        //    {
        //        #region CloseDriver
        //        s_InfoStrIntegration = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Close Driver", ScriptPara.Index, ScriptPara.ActionName);
        //        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrIntegration, txtIntegrationFunctionTestInformation });
        //        cs_BrowserIntegration.Close_WebDriver();
        //        #endregion
        //    }

        //    ---------------------------------------//
        //    --------------- OpenDriver ---------------//
        //    ---------------------------------------//
        //    else if (ScriptPara.Action.CompareTo("OpenDriver") == 0)
        //    {
        //        #region OpenDriver
        //        s_InfoStrIntegration = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Open Driver", ScriptPara.Index, ScriptPara.ActionName);
        //        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrIntegration, txtIntegrationFunctionTestInformation });
        //        TestInitialIntegration();
        //        #endregion
        //    }


        //    WriteconsoleLogIntegration();



        //    string Key1 = "Login-";
        //    string Key2 = "ReLogin";


        //    if (ScriptPara.ActionName.ToLower().IndexOf(Key1.ToLower()) < 0 && ScriptPara.ActionName.ToLower().IndexOf(Key2.ToLower()) < 0)
        //    {
        //        ---------- Write Test Report ----------//
        //        WriteTestReportIntegration(j, TestResult, ref DATA_ROW, ScriptPara.Note);
        //        WriteWebGUITestReportIntegration(bTestResult, ScriptPara);
        //    }

        //    if (ScriptPara.Note != string.Empty)
        //    {
        //        st_ReadFinalScriptDataIntegration[i_FinalScriptIndexIntegration].TestResult = "FAIL";
        //        st_ReadFinalScriptDataIntegration[i_FinalScriptIndexIntegration].Comment = "FAIL in: \n" + st_ReadStepsScriptDataIntegration[i_StepsScriptIndexIntegration].Name;
        //        i_duringStartStopIntegration = Convert.ToInt32(st_ReadFinalScriptDataIntegration[i_FinalScriptIndexIntegration].StopIndex) + 1;
        //    }
        //    bWebGUISingleScriptItemRunning = false;
        //}





        

        
    }
}
