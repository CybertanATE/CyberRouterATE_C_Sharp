
#define DEBUG_MODE
#undef DEBUG_MODE
#define GUI_COMMON_CASE

#define TIMEOUE_SETTING
//#undef TIMEOUE_SETTING



using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Xml;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Net.NetworkInformation;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using ComportClass;
using NS_CbtSeleniumApi;
using Ns_CbtFtpClient;
using NS_cbtLineNotificationApi;


namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        string s_ReportFileNameUSBStorage = string.Empty;
        string s_ReportFileNameWebGUIUSBStorage = string.Empty;
        string s_ReportPathUSBStorage = System.Windows.Forms.Application.StartupPath + @"\report";
        string s_ReportFilePathUSBStorage;
        string s_ReportFilePathWebGUIUSBStorage;
        string s_ConsoleTextUSBStorage = string.Empty;
        string s_InfoStrUSBStorage;
        string s_BrowserDriverProcessesNameUSBStorage = string.Empty;
        string s_FolderNameUSBStorage = "EA8500USBStorage_DATE";
        string s_LogFileUSBStorage = string.Empty;

        bool bWebGUISingleScriptItemRunning = false;
        Thread thread_WebGUIScriptItem;
        CBT_SeleniumApi.BrowserType csbt_TestBrowserUSBStorage;
        //string sTestBrowser = string.Empty;
        List<CBT_SeleniumApi.BrowserType> csbt_TestBrowserListUSBStorage = new List<CBT_SeleniumApi.BrowserType>();
        //List<string> sTestBrowserList = new List<string>();
        ExcelObject st_ExcelObjectUSBStorage = new ExcelObject();

        ModelGroupStruct mgs_ModelParameterUSBStorage = new ModelGroupStruct();
        LoginSettingGroupStruct lsgs_GatewayLoginSettingParameterUSBStorage = new LoginSettingGroupStruct();
        LoginSetting ls_LoginSettingParametersUSBStorage = new LoginSetting();
        TestBrowerSetting tbs_TestBrowerParametersUSBStorage = new TestBrowerSetting();

        int i_TestRun = 1;
        int i_FinalScriptIndexUSBStorage = 0;
        int i_StepsScriptIndexUSBStorage = 0;
        int i_duringStartStopUSBStorage = 0;
        int i_TestScriptIndexUSBStorage = 0;
        int i_DATA_FinalROWUSBStorage = 16;
        int i_DATA_ROWUSBStorage = 15;
        int i_RetryCountUSBStorage = 1;
        int i_CommandDataIndex = 0;
        string s_CurrentURLUSBStorage = string.Empty;
        Stopwatch sw_TestTimerUSBStorage = new Stopwatch();
        Stopwatch sw_WebGUITestTimerUSBStorage = new Stopwatch();

        CBT_SeleniumApi cs_BrowserUSBStorage = null;
        Comport cp_ComPortUSBStorage = null;
        CbtFtpClient ftpclient = null;

        List<CbtLineNotificationApi> cln_LineNotificationListUSBStorage = new List<CbtLineNotificationApi>();


        FinalScriptDataUSBStorage[] st_ReadFinalScriptDataUSBStorage;
        StepsScriptDataUSBStorage[] st_ReadStepsScriptDataUSBStorage;
        ScriptDataUSBStorage[] st_ReadScriptDataUSBStorage;

        //**********************************************************************************//
        //------------------ USB Storage Function Test Event -------------------//
        //**********************************************************************************//
        #region USB Storage Test Event
        private void USBStorageTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //TestItem = TestItemConstants.TESTITEM_THROUGHPUT;
            ToggleToolStripMenuItem(false);
            tabControl_RouterStartPage.Hide();

            USBStorageTestToolStripMenuItem.Checked = true;

            HideAllTabPage();

            this.tpUSBStorageFunctionTest.Parent = this.tabControl_USBStorage;
            tabControl_USBStorage.Show();
            tsslMessage.Text = tabControl_USBStorage.TabPages[tabControl_USBStorage.SelectedIndex].Text + " Control Panel";

            txtUSBStorageFunctionTestFinalScriptFilePath.Text = System.Windows.Forms.Application.StartupPath + "\\testCondition\\USBStorageFinal.xlsm";
            txtUSBStorageFunctionTestStepsScriptFilePath.Text = System.Windows.Forms.Application.StartupPath + "\\testCondition\\USBStorageSteps.xlsx";
            txtUSBStorageFunctionTestScriptFilePath.Text = System.Windows.Forms.Application.StartupPath + "\\testCondition\\DutWebScript\\EA8500_WebGuiScript.xlsx";

            string xmlFile = System.Windows.Forms.Application.StartupPath + "\\config\\" + Constants.XML_ROUTER_USBStorage;
            Debug.WriteLine(File.Exists(xmlFile) ? "File exist" : "File is not existing");
        }

        private void tabControl_USBStorage_Selected(object sender, TabControlEventArgs e)
        {
            //ElementVisibleSetting();
            if (tabControl_USBStorage.SelectedIndex >= 0)
            {
                tsslMessage.Text = tabControl_USBStorage.TabPages[tabControl_USBStorage.SelectedIndex].Text + " Control Panel";
            }
        }

        private void btnUSBStorageFunctionTestFinalScriptBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogFinal = new OpenFileDialog();

            openFileDialogFinal.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            openFileDialogFinal.Filter = "Excel file|*.xlsx|Excel file|*.xlsm";

            if (openFileDialogFinal.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialogFinal.FileName != "")
            {
                txtUSBStorageFunctionTestFinalScriptFilePath.Clear();
                txtUSBStorageFunctionTestFinalScriptFilePath.Text = openFileDialogFinal.FileName;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            }
            else
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                MessageBox.Show("Please Select a Script File!!");
                return;
            }
            //MessageBox.Show(new Form { TopMost = true }, "Load Final Script Successfully.", "ATE Information", MessageBoxButtons.OK);
        }

        private void btnUSBStorageFunctionTestStepsScriptBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogSteps = new OpenFileDialog();

            openFileDialogSteps.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            openFileDialogSteps.Filter = "Excel file|*.xlsx";

            if (openFileDialogSteps.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialogSteps.FileName != "")
            {
                txtUSBStorageFunctionTestStepsScriptFilePath.Clear();
                txtUSBStorageFunctionTestStepsScriptFilePath.Text = openFileDialogSteps.FileName;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            }
            else
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                MessageBox.Show("Please Select a Script File!!");
                return;
            }
            //MessageBox.Show(new Form { TopMost = true }, "Load Steps Script Successfully.", "ATE Information", MessageBoxButtons.OK);
        }

        private void btnUSBStorageFunctionTestScriptBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogWebGUI = new OpenFileDialog();

            openFileDialogWebGUI.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            // Set filter for file extension and default file extension
            openFileDialogWebGUI.Filter = "Excel file|*.xlsx";

            // If the file name is not an empty string open it for opening.
            if (openFileDialogWebGUI.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialogWebGUI.FileName != "")
            {
                txtUSBStorageFunctionTestScriptFilePath.Clear();
                txtUSBStorageFunctionTestScriptFilePath.Text = openFileDialogWebGUI.FileName;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            }
            else
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                MessageBox.Show("Please Select a Script File!!");
                return;
            }
            //MessageBox.Show(new Form { TopMost = true }, "Load Script Successfully.", "ATE Information", MessageBoxButtons.OK);
        }

        private void btnUSBStorageFunctionTestSave_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Save Btn Checked!!");
            SaveFileDialog saveDebugFileDialog = new SaveFileDialog();

            saveDebugFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            saveDebugFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveDebugFileDialog.FilterIndex = 1;
            saveDebugFileDialog.RestoreDirectory = true;
            saveDebugFileDialog.FileName = "TestLog";

            if (saveDebugFileDialog.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText(saveDebugFileDialog.FileName, txtUSBStorageFunctionTestInformation.Text);
            }
        }

        private void btnUSBStorageFunctionTestRun_Click(object sender, EventArgs e)
        {
            btnUSBStorageFunctionTestRun.Enabled = false;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            if (bCommonFTThreadRunning == false && bCommonTestComplete == true)
            {
                SaveModelParameterUSBStorage();
                SaveGatewayLoginParameterUSBStorage();
                SaveTestBrowserParameterUSBStorage();
                USBStorageFunctionTestLineEnable();
                if (!CheckNeededParameterUSBStorage())
                {
                    Thread.Sleep(1000);
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    btnUSBStorageFunctionTestRun.Enabled = true;
                    return;
                }


                //-------------------------------------------------//
                //----------------- Start Test --------------------//
                //-------------------------------------------------//

                bCommonFTThreadRunning = true;
                bCommonTestComplete = false;
                ToggleFunctionTestControllerUSBStorage(false);  /* Disable all controller */
                txtUSBStorageFunctionTestInformation.Clear();

                //--------------- Start Test Thread ---------------//
                SetText("-------------- Start USBStorage ATE Test --------------", txtUSBStorageFunctionTestInformation);
                threadCommonFT = new Thread(new ThreadStart(DoFunctionTestUSBStorage));
                threadCommonFT.Name = "";
                threadCommonFT.Start();

                //*** Use another thread to catch the stop event of test thread ***//
                threadCommonFTstopEvent = new Thread(new ThreadStart(threadFunctionTestCatchStopEventUSBStorage));
                threadCommonFTstopEvent.Name = "";
                threadCommonFTstopEvent.Start();
            }
            else if (bCommonFTThreadRunning == true && bCommonTestComplete == false)
            {
                tsslMessage.Text = "Function Test Control Panel";
                bCommonFTThreadRunning = false;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "The Test will stop later. Please wait!", "ATE Information", MessageBoxButtons.OK);
            }

            /* Button release */
            Thread.Sleep(1000);
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            btnUSBStorageFunctionTestRun.Enabled = true;  //也讓Stop可以點選

        }


        #endregion

        //**********************************************************************************//
        //------------------ USB Storage Function Test Module ------------------//
        //**********************************************************************************//
        #region USB Storage Test Module


        private bool CheckNeededParameterUSBStorage()
        {
            if (masktxtUSBStorageFunctionTestGatewayIP.Text == "" || masktxtUSBStorageFunctionTestGatewayIP.Text == null)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Set IP!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (nudUSBStorageFunctionTestDefaultSettingsHTTPport.Text == "" || nudUSBStorageFunctionTestDefaultSettingsHTTPport.Text == null)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Set HTTP port!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (txtUSBStorageFunctionTestFinalScriptFilePath.Text == "" || txtUSBStorageFunctionTestFinalScriptFilePath.Text == null)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Select the Final Script!!", "Warning", MessageBoxButtons.OK);
                return false;
            }
            else if (!System.IO.File.Exists(txtUSBStorageFunctionTestFinalScriptFilePath.Text))
            {
                MessageBox.Show(new Form { TopMost = true }, "Final Script file does not exist!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (txtUSBStorageFunctionTestStepsScriptFilePath.Text == "" || txtUSBStorageFunctionTestStepsScriptFilePath.Text == null)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Select the Steps Script!!", "Warning", MessageBoxButtons.OK);
                return false;
            }
            else if (!System.IO.File.Exists(txtUSBStorageFunctionTestStepsScriptFilePath.Text))
            {
                MessageBox.Show(new Form { TopMost = true }, "Steps Script file does not exist!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (txtUSBStorageFunctionTestScriptFilePath.Text == "" || txtUSBStorageFunctionTestScriptFilePath.Text == null)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Select the Test Script!!", "Warning", MessageBoxButtons.OK);
                return false;
            }
            else if (!System.IO.File.Exists(txtUSBStorageFunctionTestScriptFilePath.Text))
            {
                MessageBox.Show(new Form { TopMost = true }, "WebGUI Script file does not exist!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (!CheckTestBrowserUSBStorage())
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Select the Test Browser!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (ckboxUSBStorageFunctionTestConsoleLog.Checked == true)
            {
                cp_ComPortUSBStorage = new Comport();

                if (cp_ComPortUSBStorage.isOpen() == false)
                {
                    MessageBox.Show(new Form { TopMost = true }, "COM Port is not Open!!", "Warning", MessageBoxButtons.OK);
                    return false;
                }
            }
            return true;
        }

        private void threadFunctionTestCatchStopEventUSBStorage()
        {
            while (bCommonFTThreadRunning == true)
            {
                //Thread.Sleep(500);
            }

            TestThreadFinishedUSBStorage();
            threadCommonFTstopEvent.Abort();
        }

        private void TestThreadFinishedUSBStorage()
        {
            foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
            {
                cln_LineNotificationUSBStorage.postMessage("---USB Storage ATE Test End---");
                cln_LineNotificationUSBStorage.postMessageAndSticker("Congratulations!", 2, 144);
                cln_LineNotificationUSBStorage.postMessageAndPicture("by CyberTan", System.Windows.Forms.Application.StartupPath + @"\images.jpg");
            }
            //cln_LineNotificationUSBStorage2.postMessage("---USB Storage ATE Test End---");
            //cln_LineNotificationUSBStorage3.postMessage("---USB Storage ATE Test End---");
            //cln_LineNotificationUSBStorage4.postMessage("---USB Storage ATE Test End---");
            //cln_LineNotificationUSBStorage5.postMessage("---USB Storage ATE Test End---");
            //cln_LineNotificationUSBStorage6.postMessage("---USB Storage ATE Test End---");
            this.Invoke(new showCommonGUIDelegate(ToggleFunctionTestUSBStorage));
            Invoke(new SetTextCallBack(SetText), new object[] { "--------------- EA8500 USB Storage ATE Test End--------------", txtUSBStorageFunctionTestInformation });
        }

        private void DoFunctionTestUSBStorage()
        {
            WriteDebugMsg = new System.IO.StreamWriter(sDebugFilePath);
            WriteConsoleLog = new System.IO.StreamWriter(sConsoleLogFilePath);
            int iTestRun = Convert.ToInt32(tbs_TestBrowerParametersUSBStorage.TestRun);
            int iTestTimeOut; // = Convert.ToInt32(tbsTestBrowerParameters.TestTimeOut) * 1000;

            s_InfoStrUSBStorage = string.Format("Loading Final Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
            loadFinalScript();
            s_InfoStrUSBStorage = string.Format("Loading Steps Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
            loadStepsScript();
            s_InfoStrUSBStorage = string.Format("Loading WebGUI Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
            loadWebGUIScript();
            for (i_TestRun = 1; i_TestRun <= iTestRun; i_TestRun++)
            {
                foreach (CBT_SeleniumApi.BrowserType csbtBrowserStr in csbt_TestBrowserListUSBStorage)
                {
                    i_TestScriptIndexUSBStorage = 0;
                    i_DATA_FinalROWUSBStorage = 16;
                    i_DATA_ROWUSBStorage = 15;
                    s_CurrentURLUSBStorage = string.Empty;

                    csbt_TestBrowserUSBStorage = csbtBrowserStr;
                    s_FolderNameUSBStorage = "EA8500USBStorage_DATE";
                    CreateNewExcelFileUSBStorage();
                    CreateNewWebGUIExcelFileUSBStorage();

                    s_InfoStrUSBStorage = string.Format("<< Test by Bowser: {0}   Test Run:{1} >>\n", csbtBrowserStr, i_TestRun);
                    Invoke(new SetTextCallBack(SetTextC), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });

                    if (ckboxUSBStorageFunctionTestConsoleLog.Checked == true)
                    {
                        s_InfoStrUSBStorage = string.Format("------------------------------------- << Test by Bowser: {0}   Test Run:{1} >> -------------------------------------", csbtBrowserStr, i_TestRun);
                        WriteLogTitleUSBStorage(WriteConsoleLog, s_InfoStrUSBStorage, "=");
                    }
                    for (i_FinalScriptIndexUSBStorage = 0; i_FinalScriptIndexUSBStorage < st_ReadFinalScriptDataUSBStorage.Length; i_FinalScriptIndexUSBStorage++)
                    {
                        //tempLogColumn = 8;
                        i_RetryCountUSBStorage = 1;

                        s_LogFileUSBStorage = s_ReportPathUSBStorage + @"\" + s_FolderNameUSBStorage + @"\" + "ConsoleLog_Index" + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestItem + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";

                        s_InfoStrUSBStorage = string.Format("========Index {0} Start=========\n", st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestItem);
                        Invoke(new SetTextCallBack(SetTextC), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                        s_InfoStrUSBStorage = string.Format("Function Name: {0}, Name: {1}", st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].FunctionName, st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name);
                        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                        s_InfoStrUSBStorage = string.Format("Sub-Index: Start: {0}, Stop: {1}", st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StartIndex, st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex);
                        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });

                        for (i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StartIndex); i_duringStartStopUSBStorage <= Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex); i_duringStartStopUSBStorage++)
                        {
                            for (i_StepsScriptIndexUSBStorage = 0; i_StepsScriptIndexUSBStorage < st_ReadStepsScriptDataUSBStorage.Length; i_StepsScriptIndexUSBStorage++)
                            {
                                if (Convert.ToInt32(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index) == i_duringStartStopUSBStorage)
                                {
                                    s_InfoStrUSBStorage = string.Format("\r\n=== Run Step-Index {0} Start===", st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index);
                                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });

                                    //----------------------------------------------------------------//
                                    //----------------------- Main Test Thread -----------------------//
                                    //----------------------------------------------------------------//
                                    sw_TestTimerUSBStorage.Reset();
                                    sw_TestTimerUSBStorage.Start();
                                    bCommonSingleScriptItemRunning = true;
                                    threadCommonFTscriptItem = new Thread(new ThreadStart(FunctionTestMainFunctionUSBStorage));
                                    //threadCommonFTscriptItem.Name = "";
                                    threadCommonFTscriptItem.Start();
                                    int timeoutNumber = 0;
                                    if (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FunctionName == "WebGUI")
                                    {
                                        iTestTimeOut = 20 * 60 * 1000;
                                    }
                                    else if (Int32.TryParse(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].TestTimeOut, out timeoutNumber) == true)
                                    {
                                        iTestTimeOut = Convert.ToInt32(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].TestTimeOut) * 1000;
                                    }
                                    else
                                    {
                                        iTestTimeOut = Convert.ToInt32(nudUSBStorageFunctionTestDefaultTimeout.Value) * 1000;
                                    }
                                    while (sw_TestTimerUSBStorage.ElapsedMilliseconds <= iTestTimeOut && bCommonSingleScriptItemRunning == true)
                                    {
                                        //Thread.Sleep(500);
                                    }
                                    sw_TestTimerUSBStorage.Stop();
                                    string ElapsedTestTime = sw_TestTimerUSBStorage.ElapsedMilliseconds.ToString();


                                    //----------------------------------------------//
                                    //------------------ Time Out ------------------//
                                    //----------------------------------------------//
                                    #region Time Out

                                    if (sw_TestTimerUSBStorage.ElapsedMilliseconds > iTestTimeOut)
                                    {
                                        try
                                        {
                                            thread_WebGUIScriptItem.Abort();
                                        }
                                        catch { }

                                        threadCommonFTscriptItem.Abort();
                                        st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "FAIL";


                                    }
                                    #endregion
                                }
                            }
                        }
                        if (st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult == "")
                        {
                            st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "PASS";
                        }
                        WriteTestReportUSBStorage(true, st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage]);
                        File.AppendAllText(s_LogFileUSBStorage, txtUSBStorageFunctionTestInformation.Text);
                        try
                        {
                            cs_BrowserUSBStorage.Close_WebDriver();
                        }
                        catch { }
                    }
                    TestFinishedWebGUIActionUSBStorage();
                    TestFinishedActionUSBStorage();
                }
            }

            bCommonFTThreadRunning = false;
        }

        private void WebGUIMainFunctionUSBStorage()
        {
            Thread.Sleep(2000);
            CBT_SeleniumApi.GuiScriptParameter ScriptPara = new CBT_SeleniumApi.GuiScriptParameter();
            string sExceptionInfo = string.Empty;
            string sLoginURL = string.Format("http://{0}:{1}", ls_LoginSettingParametersUSBStorage.GatewayIP, ls_LoginSettingParametersUSBStorage.HTTP_Port);
            //string sCurrentURL = string.Empty;
            //int DATA_ROW = 15;
            bool bTestResult = true;
            int j = i_TestScriptIndexUSBStorage;

            ScriptPara.Procedure = st_ReadScriptDataUSBStorage[j].Procedure;
            ScriptPara.Index = st_ReadScriptDataUSBStorage[j].TestIndex;
            ScriptPara.Action = st_ReadScriptDataUSBStorage[j].Action;
            ScriptPara.ActionName = st_ReadScriptDataUSBStorage[j].ActionName;
            ScriptPara.ElementType = st_ReadScriptDataUSBStorage[j].ElementType;
            ScriptPara.ElementXpath = st_ReadScriptDataUSBStorage[j].ElementXpath;
            ScriptPara.ElementXpath = ScriptPara.ElementXpath.Replace('\"', '\'');
            ScriptPara.RadioBtnExpectedValueXpath = st_ReadScriptDataUSBStorage[j].RadioButtonExpectedValueXpath.Replace('\"', '\'');
            ScriptPara.WriteValue = st_ReadScriptDataUSBStorage[j].WriteExpectedValue;
            ScriptPara.ExpectedValue = st_ReadScriptDataUSBStorage[j].WriteExpectedValue;
            ScriptPara.URL = sLoginURL + st_ReadScriptDataUSBStorage[j].WriteExpectedValue;
            ScriptPara.TestTimeOut = st_ReadScriptDataUSBStorage[j].TestTimeOut;
            ScriptPara.Note = string.Empty;

            //---------------------------------------//
            //-------------- Go To URL --------------//
            //---------------------------------------//
            if (j == 0)
            {
                s_CurrentURLUSBStorage = sLoginURL;
                cs_BrowserUSBStorage.GoToURL(sLoginURL);
                Thread.Sleep(1000);
            }
            if (ScriptPara.Action.CompareTo("Goto") == 0 && s_CurrentURLUSBStorage.CompareTo(ScriptPara.URL) != 0)
            {
                #region Go To URL
                s_CurrentURLUSBStorage = ScriptPara.URL;
                Thread.Sleep(1000);
                cs_BrowserUSBStorage.GoToURL(s_CurrentURLUSBStorage);
                #endregion
            }

            //---------------------------------------//
            //-------------- Set Value --------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("Set") == 0)
            {
                #region Set Value
                s_InfoStrUSBStorage = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Set Value", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });

                sw_TestTimerUSBStorage.Stop();
                sw_WebGUITestTimerUSBStorage.Stop();
                bool checkXPathResult = false;
                Stopwatch waitingXpath = new Stopwatch();
                waitingXpath.Start();
                do
                {
                    checkXPathResult = cs_BrowserUSBStorage.CheckXPathDisplayed(ScriptPara.ElementXpath);
                } while (waitingXpath.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);
                waitingXpath.Reset();
                sw_TestTimerUSBStorage.Start();
                sw_WebGUITestTimerUSBStorage.Start();

                if (checkXPathResult == true)
                {
                    Thread.Sleep(2000);
                    try
                    {           //So that the StepsScript could set parameters such as "test1, test2...." 
                        if (ScriptPara.WriteValue != "" && st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].WriteValue[i_CommandDataIndex] != "" && st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].WriteValue[i_CommandDataIndex] != null)
                        {
                            ScriptPara.WriteValue = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].WriteValue[i_CommandDataIndex];
                            i_CommandDataIndex++;
                        }
                    }
                    catch { }
                    ScriptPara.Note = string.Empty;
                    try
                    {
                        cs_BrowserUSBStorage.SetWebElementValue(ref ScriptPara);
                    }
                    catch
                    {
                        ExceptionActionUSBStorage(s_CurrentURLUSBStorage, ref ScriptPara);
                        ScriptPara.Note = "Set Value Error:\n" + ScriptPara.Note;
                        SwitchToRGconsoleUSBStorage();
                        WriteconsoleLogUSBStorage();
                        //return false;
                    }
                }
                else if (checkXPathResult == false)
                {
                    s_InfoStrUSBStorage = string.Format("...Couldn't find the element!");
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                    ExceptionActionUSBStorage(s_CurrentURLUSBStorage, ref ScriptPara);
                    SwitchToRGconsoleUSBStorage();
                    WriteconsoleLogUSBStorage();
                }


                #endregion
            }

            //---------------------------------------//
            //-------------- Get Value --------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("Get") == 0)
            {
                #region Get Value
                s_InfoStrUSBStorage = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Get Value", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });

                ScriptPara.Note = string.Empty;
                ScriptPara.GetValue = string.Empty;

                if (ScriptPara.ElementType.CompareTo("RADIO_BUTTON") == 0)
                {
                    ScriptPara.ElementXpath = st_ReadScriptDataUSBStorage[j].RadioButtonExpectedValueXpath.Replace('\"', '\'');
                }

                try
                {
                    bTestResult = cs_BrowserUSBStorage.GetWebElementValue(ref ScriptPara); // Set Value
                }
                catch
                {
                    ExceptionActionUSBStorage(s_CurrentURLUSBStorage, ref ScriptPara);
                    ScriptPara.Note = "Get value Error:\n" + ScriptPara.Note;
                    //return false;
                }


                //---------- Write Test Report ----------//
                //WriteTestReportUSBStorage(j, TestResult, ref DATA_ROW, ScriptPara.Note);
                #endregion
            }

            //---------------------------------------//
            //--------------- ReLogin ---------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("ReLogin") == 0)
            {
                #region ReLogin
                Invoke(new SetTextCallBack(SetText), new object[] { "Log in again...", txtUSBStorageFunctionTestInformation });
                Re_LoginUSBStorage();
                #endregion
            }

            //---------------------------------------//
            //--------------- Alert Login ---------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("AlertLogin") == 0)
            {
                #region Alert Login
                Stopwatch pingTimer = new Stopwatch();
                pingTimer.Reset();
                pingTimer.Start();
                Invoke(new SetTextCallBack(SetText), new object[] { "Login...", txtUSBStorageFunctionTestInformation });
                while (true)
                {
                    if (pingTimer.ElapsedMilliseconds > (Convert.ToInt32(120000)))
                    {
                        break;
                    }

                    if (PingClient(ls_LoginSettingParametersUSBStorage.GatewayIP, 1000))
                    {
                        Invoke(new SetTextCallBack(SetText), new object[] { "\tPing Successfully!!", txtUSBStorageFunctionTestInformation });
                        break;
                    }

                    Thread.Sleep(1000);
                    string Info = String.Format(".");
                    Invoke(new SetTextCallBack(SetText), new object[] { Info, txtUSBStorageFunctionTestInformation });
                }
                Thread.Sleep(3000);
                String[] splitLoginInfo = ScriptPara.WriteValue.Split('/');
                cs_BrowserUSBStorage.loginAlertMessage(ScriptPara.URL, splitLoginInfo[0], splitLoginInfo[1]);
                Thread.Sleep(3000);
                //SendKeys.Send(" {Enter}");
                #endregion
            }

            //---------------------------------------//
            //-------------- File Upload --------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("FileUpload") == 0)
            {
                #region File Upload
                s_InfoStrUSBStorage = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:File Upload", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });

                Thread.Sleep(3000);
                ScriptPara.Note = string.Empty;
                try
                {
                    cs_BrowserUSBStorage.fileUploadE8350(ref ScriptPara);
                }
                catch
                {
                    ExceptionActionUSBStorage(s_CurrentURLUSBStorage, ref ScriptPara);
                    ScriptPara.Note = "File Upload Error:\n" + ScriptPara.Note;
                    SwitchToRGconsoleUSBStorage();
                    WriteconsoleLogUSBStorage();
                }
                #endregion
            }

            //---------------------------------------//
            //--------------- Wait ---------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("Wait") == 0)
            {
                #region Wait
                ScriptPara.Note = string.Empty;
                if (ScriptPara.ElementXpath.CompareTo("") == 0)
                {
                    s_InfoStrUSBStorage = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Waiting for {2} seconds", ScriptPara.Index, ScriptPara.ActionName, ScriptPara.TestTimeOut);
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                    sw_TestTimerUSBStorage.Stop();
                    Thread.Sleep(Convert.ToInt16(ScriptPara.TestTimeOut) * 1000);
                    sw_TestTimerUSBStorage.Start();
                    SwitchToRGconsoleUSBStorage();
                    WriteconsoleLogUSBStorage();
                }
                else
                {
                    s_InfoStrUSBStorage = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Waiting for the XPath, it won't be more than {2} seconds", ScriptPara.Index, ScriptPara.ActionName, ScriptPara.TestTimeOut);
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                    sw_TestTimerUSBStorage.Stop();
                    bool checkXPathResult = false;
                    Stopwatch waitingXpath = new Stopwatch();
                    waitingXpath.Start();
                    do
                    {
                        checkXPathResult = cs_BrowserUSBStorage.CheckXPathDisplayed(ScriptPara.ElementXpath);
                    } while (waitingXpath.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);
                    waitingXpath.Reset();
                    sw_TestTimerUSBStorage.Start();
                    SwitchToRGconsoleUSBStorage();
                    WriteconsoleLogUSBStorage();
                }
                #endregion
            }

            //---------------------------------------//
            //--------------- HoldMenu --------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("HoldMenu") == 0)
            {
                #region HoldMenu
                s_InfoStrUSBStorage = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Hold Menu", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });

                Thread.Sleep(3000);
                ScriptPara.Note = string.Empty;
                try
                {
                    cs_BrowserUSBStorage.holdMenu(ref ScriptPara);
                }
                catch
                {
                    ExceptionActionUSBStorage(s_CurrentURLUSBStorage, ref ScriptPara);
                    ScriptPara.Note = "Hold Menu Error:\n" + ScriptPara.Note;
                    SwitchToRGconsoleUSBStorage();
                    WriteconsoleLogUSBStorage();
                }
                #endregion
            }

            //---------------------------------------//
            //-------------- Poke --------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("Poke") == 0)
            {
                #region Set Value
                s_InfoStrUSBStorage = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Poke", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });

                sw_TestTimerUSBStorage.Stop();
                sw_WebGUITestTimerUSBStorage.Stop();
                bool checkXPathResult = false;
                Stopwatch waitingXpath = new Stopwatch();
                waitingXpath.Start();
                do
                {
                    checkXPathResult = cs_BrowserUSBStorage.CheckXPathDisplayed(ScriptPara.ElementXpath);
                } while (waitingXpath.ElapsedMilliseconds < 10 * 1000 && checkXPathResult == false);
                waitingXpath.Reset();
                sw_TestTimerUSBStorage.Start();
                sw_WebGUITestTimerUSBStorage.Start();

                if (checkXPathResult == true)
                {
                    Thread.Sleep(2000);
                    try
                    {           //So that the StepsScript could set parameters such as "test1, test2...." 
                        if (ScriptPara.WriteValue != "" && st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].WriteValue[i_CommandDataIndex] != "" && st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].WriteValue[i_CommandDataIndex] != null)
                        {
                            ScriptPara.WriteValue = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].WriteValue[i_CommandDataIndex];
                            i_CommandDataIndex++;
                        }
                    }
                    catch { }
                    ScriptPara.Note = string.Empty;
                    try
                    {
                        cs_BrowserUSBStorage.SetWebElementValue(ref ScriptPara);
                    }
                    catch
                    {
                        s_InfoStrUSBStorage = string.Format("Didn't poke(catch)");
                        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                    }
                }
                else if (checkXPathResult == false)
                {
                    s_InfoStrUSBStorage = string.Format("Didn't poke");
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                }


                #endregion
            }

            //---------------------------------------//
            //--------------- CloseDriver ---------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("CloseDriver") == 0)
            {
                #region CloseDriver
                s_InfoStrUSBStorage = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Close Driver", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                try
                {
                    cs_BrowserUSBStorage.Close_WebDriver();
                }
                catch { }
                #endregion
            }

            //---------------------------------------//
            //--------------- OpenDriver ---------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("OpenDriver") == 0)
            {
                #region OpenDriver
                s_InfoStrUSBStorage = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Open Driver", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                TestInitialUSBStorage();
                #endregion
            }


            WriteconsoleLogUSBStorage();



            string Key1 = "Login-";
            string Key2 = "ReLogin";


            if (ScriptPara.ActionName.ToLower().IndexOf(Key1.ToLower()) < 0 && ScriptPara.ActionName.ToLower().IndexOf(Key2.ToLower()) < 0)
            {
                //---------- Write Test Report ----------//
                //WriteTestReportUSBStorage(j, TestResult, ref DATA_ROW, ScriptPara.Note);
                WriteWebGUITestReportUSBStorage(bTestResult, ScriptPara);
            }

            if (ScriptPara.Note != string.Empty)
            {
                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "FAIL";
                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "FAIL in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                s_InfoStrUSBStorage = string.Format(ScriptPara.Note);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                {
                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                }
                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
            }
            bWebGUISingleScriptItemRunning = false;
        }

        private void FunctionTestMainFunctionUSBStorage()
        {
            switch (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FunctionName)
            {
                case "WebGUI":
                    //----------------------------------------------------------------//
                    //---------------------------- WebGUI ----------------------------//
                    //----------------------------------------------------------------//
                    #region WebGUI
                    s_InfoStrUSBStorage = string.Format("** [Step Index]: {0}   [Function Name]:WebGUI    [Name]:{1} ", st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index, st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                    i_CommandDataIndex = 0;
                    for (i_TestScriptIndexUSBStorage = 0; i_TestScriptIndexUSBStorage < st_ReadScriptDataUSBStorage.Length; i_TestScriptIndexUSBStorage++)
                    {
                        if (st_ReadScriptDataUSBStorage[i_TestScriptIndexUSBStorage].Procedure == st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Action)
                        {
                            if (st_ReadScriptDataUSBStorage[i_TestScriptIndexUSBStorage].TestIndex.CompareTo("") != 0 && st_ReadScriptDataUSBStorage[i_TestScriptIndexUSBStorage].Action.CompareTo("") != 0)
                            {
                                bWebGUISingleScriptItemRunning = true;
                                int iWebGUITestTimeOut = 120 * 1000;
                                if (i_TestScriptIndexUSBStorage == 0)
                                {
                                    TestInitialUSBStorage();
                                }
                                if (bWebGUISingleScriptItemRunning == true)
                                {
                                    sw_WebGUITestTimerUSBStorage.Reset();
                                    sw_WebGUITestTimerUSBStorage.Start();
                                    thread_WebGUIScriptItem = new Thread(new ThreadStart(WebGUIMainFunctionUSBStorage));
                                    thread_WebGUIScriptItem.Start();
                                    int iTimeOutNull = 0;
                                    if (Int32.TryParse(st_ReadScriptDataUSBStorage[i_TestScriptIndexUSBStorage].TestTimeOut, out iTimeOutNull) == true)
                                    {
                                        iWebGUITestTimeOut = Convert.ToInt32(st_ReadScriptDataUSBStorage[i_TestScriptIndexUSBStorage].TestTimeOut) * 1000;
                                    }
                                    while (sw_WebGUITestTimerUSBStorage.ElapsedMilliseconds <= iWebGUITestTimeOut && bWebGUISingleScriptItemRunning == true)
                                    {
                                        //Thread.Sleep(500);
                                    }
                                    sw_WebGUITestTimerUSBStorage.Stop();


                                    //----------------------------------------------//
                                    //--------------- WebGUI Time Out --------------//
                                    //----------------------------------------------//
                                    #region Time Out
                                    if (sw_WebGUITestTimerUSBStorage.ElapsedMilliseconds > iWebGUITestTimeOut)
                                    {
                                        thread_WebGUIScriptItem.Abort();
                                        bWebGUISingleScriptItemRunning = false;

                                        string TestStatus = string.Empty;
                                        if (!CheckExtraTestStatusUSBStorage(ref TestStatus))
                                        {
                                            TestStatus = "Time Out";
                                        }
                                        else
                                        {
                                            TestStatus = string.Format("Time Out. " + TestStatus);
                                        }

                                        WriteWebGUITestReportUSBStorage(i_TestScriptIndexUSBStorage, false, TestStatus);

                                        if (StopWhenTestErrorUSBStorage())
                                        {
                                            TestFinishedWebGUIActionUSBStorage();
                                            //TestFinishedActionUSBStorage();
                                            Invoke(new SetTextCallBack(SetText), new object[] { "Test Abort.", txtUSBStorageFunctionTestInformation });
                                            bCommonFTThreadRunning = false;
                                            return;
                                        }
                                        else
                                        {
                                            if (i_RetryCountUSBStorage >= 1)
                                            {
                                                //------ Stop Item Test and Close Web Driver------//
                                                cs_BrowserUSBStorage.Close_WebDriver();
                                                Thread.Sleep(5000);

                                                //-------- Initial and Open New Web Driver--------//
                                                //TestInitialUSBStorage();
                                                string Info = string.Format("{0}!! Login again...", TestStatus);
                                                Invoke(new SetTextCallBack(SetText), new object[] { Info, txtUSBStorageFunctionTestInformation });
                                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StartIndex);
                                                i_StepsScriptIndexUSBStorage = -1;
                                                bCommonSingleScriptItemRunning = false;
                                                i_RetryCountUSBStorage--;
                                                return;
                                                //Re_LoginUSBStorage();
                                            }
                                            else
                                            {
                                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "ERROR";
                                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "ERROR in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                                {
                                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                                }
                                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;

                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    threadCommonFTscriptItem.Abort();
                                    threadCommonFT.Abort();
                                }
                                    #endregion
                                //WebGUIMainFunctionUSBStorage();
                            }
                        }
                        if (st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult == "FAIL" || st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult == "ERROR")
                        {
                            break;
                        }
                    }

                    break;
                    #endregion

                case "CMD":
                    //----------------------------------------------------------------//
                    //------------------------------ CMD -----------------------------//
                    //----------------------------------------------------------------//
                    #region CMD for Samba
                    s_InfoStrUSBStorage = string.Format("** [Step Index]: {0}   [Function Name]:CMD    [Name]:{1} ", st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index, st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
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
                    int iCommandDataIndex = 0;
                    process.StartInfo = startInfo;
                    process.Start();
                    while (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Command[iCommandDataIndex] != "" && st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Command[iCommandDataIndex] != null)
                    {
                        process.StandardInput.WriteLine(@st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Command[iCommandDataIndex]);
                        iCommandDataIndex++;
                    }
                    process.StandardInput.WriteLine("exit");
                    // Read the standard output of the process.
                    string output;
                    //string tempoutput = "ExpectedValue::";
                    while ((output = process.StandardOutput.ReadLine()) != null || (output = process.StandardError.ReadLine()) != null)
                    {
                        //tempoutput += output + "\n";
                        st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue += output + "\n";  //需改成 ].GetValue 
                        Invoke(new SetTextCallBack(SetText), new object[] { output, txtUSBStorageFunctionTestInformation });
                    }
                    //st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, tempLogColumn] = tempoutput;
                    //tempLogColumn++;
                    //需Uncomment
                    if (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue.Contains(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].ExpectedValue) == false)
                    {
                        st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "FAIL";
                        st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "FAIL in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                        foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                        {
                            cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                        }
                        i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                    }
                    process.WaitForExit();
                    process.Close();
                    break;
                    #endregion

                case "FTP":
                    //----------------------------------------------------------------//
                    //------------------------------ FTP -----------------------------//
                    //----------------------------------------------------------------//
                    #region FTP
                    bool response = false;
                    string exceptionInfo = "";
                    string sDestination = "";
                    switch (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Action)
                    {
                        case "Login":
                            #region Login
                            s_InfoStrUSBStorage = string.Format("** [Step Index]: {0}   [Function Name]:FTP    [Name]:{1} ", st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index, st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                            string sDir = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpIP;
                            string sUsername = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpUsername;
                            string sPassword = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpPassword;
                            response = false;
                            try
                            {
                                ftpclient = new CbtFtpClient(sDir, sUsername, sPassword);
                                if (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpPassiveMode == null || st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpPassiveMode.ToLower() != "false")
                                {
                                    ftpclient.passiveMode = true;
                                }
                                //int ftpServerPort = 21;
                                //if (Int32.TryParse(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpServerPort, out ftpServerPort) == true)
                                //{
                                //    ftpclient.ftpServerPort = ftpServerPort;
                                //}
                                response = ftpclient.Login(ref exceptionInfo);
                                //s_InfoStrUSBStorage = string.Format("{0}", response);
                                //Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                                if (response)
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Login successfully";
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Login successfully", txtUSBStorageFunctionTestInformation });
                                }
                                else
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Login fail: " + exceptionInfo;
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Login fail: " + exceptionInfo, txtUSBStorageFunctionTestInformation });
                                }
                            }
                            catch (Exception ex)
                            {
                                s_InfoStrUSBStorage = string.Format("**The step error in: {0}", ex);
                                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                                st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = s_InfoStrUSBStorage;
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "ERROR";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "ERROR in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            if (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue.ToLower().Contains(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].ExpectedValue.ToLower()) == false)
                            {
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "FAIL";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "FAIL in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            break;
                            #endregion

                        case "Upload":
                            #region Upload
                            s_InfoStrUSBStorage = string.Format("** [Step Index]: {0}   [Function Name]:FTP    [Name]:{1} ", st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index, st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                            sDestination = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpServerFolder;
                            response = false;
                            try
                            {
                                response = ftpclient.UploadFileUSBStorage(sDestination, ref exceptionInfo);
                                if (response)
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Upload successfully";
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Upload successfully", txtUSBStorageFunctionTestInformation });
                                }
                                else
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Upload fail: " + exceptionInfo;
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Upload fail: " + exceptionInfo, txtUSBStorageFunctionTestInformation });
                                }
                            }
                            catch (Exception ex)
                            {
                                s_InfoStrUSBStorage = string.Format("**The step error in: {0}", ex);
                                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                                st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = s_InfoStrUSBStorage;
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "ERROR";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "ERROR in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            if (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue.ToLower().Contains(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].ExpectedValue.ToLower()) == false)
                            {
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "FAIL";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "FAIL in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            break;
                            #endregion

                        case "Download":
                            #region Download
                            s_InfoStrUSBStorage = string.Format("** [Step Index]: {0}   [Function Name]:FTP    [Name]:{1} ", st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index, st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                            string sSource = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpServerFolder;
                            response = false;
                            try
                            {
                                response = ftpclient.DownloadFileUSBStorage(sSource, ref exceptionInfo);
                                if (response)
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Download successfully";
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Download successfully", txtUSBStorageFunctionTestInformation });
                                }
                                else
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Download fail: " + exceptionInfo;
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Download fail: " + exceptionInfo, txtUSBStorageFunctionTestInformation });
                                }
                            }
                            catch (Exception ex)
                            {
                                s_InfoStrUSBStorage = string.Format("**The step error in: {0}", ex);
                                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                                st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = s_InfoStrUSBStorage;
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "ERROR";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "ERROR in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            if (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue.ToLower().Contains(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].ExpectedValue.ToLower()) == false)
                            {
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "FAIL";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "FAIL in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            break;
                            #endregion

                        case "CreateFolder":
                            #region CreateFolder
                            s_InfoStrUSBStorage = string.Format("** [Step Index]: {0}   [Function Name]:FTP    [Name]:{1} ", st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index, st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                            sDestination = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpServerFolder;
                            response = false;
                            try
                            {
                                response = ftpclient.CreateFolderUSBStorage(sDestination, ref exceptionInfo);
                                if (response)
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Create folder successfully";
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Create folder successfully", txtUSBStorageFunctionTestInformation });
                                }
                                else
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Create folder fail: " + exceptionInfo;
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Create folder fail: " + exceptionInfo, txtUSBStorageFunctionTestInformation });
                                }
                            }
                            catch (Exception ex)
                            {
                                s_InfoStrUSBStorage = string.Format("**The step error in: {0}", ex);
                                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                                st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = s_InfoStrUSBStorage;
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "ERROR";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "ERROR in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            if (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue.ToLower().Contains(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].ExpectedValue.ToLower()) == false)
                            {
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "FAIL";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "FAIL in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            break;
                            #endregion

                        case "DeleteFolder":
                            #region DeleteFolder
                            s_InfoStrUSBStorage = string.Format("** [Step Index]: {0}   [Function Name]:FTP    [Name]:{1} ", st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index, st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                            sDestination = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpServerFolder;
                            response = false;
                            try
                            {
                                response = ftpclient.DeleteFolderUSBStorage(sDestination, ref exceptionInfo);
                                if (response)
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Delete folder successfully";
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Delete folder successfully", txtUSBStorageFunctionTestInformation });
                                }
                                else
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Delete folder fail: " + exceptionInfo;
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Delete folder fail: " + exceptionInfo, txtUSBStorageFunctionTestInformation });
                                }
                            }
                            catch (Exception ex)
                            {
                                s_InfoStrUSBStorage = string.Format("**The step error in: {0}", ex);
                                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                                st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = s_InfoStrUSBStorage;
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "ERROR";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "ERROR in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            if (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue.ToLower().Contains(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].ExpectedValue.ToLower()) == false)
                            {
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "FAIL";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "FAIL in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            break;
                            #endregion

                        case "Rename":
                            #region Rename
                            s_InfoStrUSBStorage = string.Format("** [Step Index]: {0}   [Function Name]:FTP    [Name]:{1} ", st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index, st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                            sDestination = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpServerFolder;
                            response = false;
                            try
                            {
                                response = ftpclient.RenameUSBStorage(sDestination, st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpRenameTo, ref exceptionInfo);
                                if (response)
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Rename successfully";
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Rename successfully", txtUSBStorageFunctionTestInformation });
                                }
                                else
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "Rename fail: " + exceptionInfo;
                                    Invoke(new SetTextCallBack(SetText), new object[] { "Rename fail: " + exceptionInfo, txtUSBStorageFunctionTestInformation });
                                }
                            }
                            catch (Exception ex)
                            {
                                s_InfoStrUSBStorage = string.Format("**The step error in: {0}", ex);
                                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                                st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = s_InfoStrUSBStorage;
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "ERROR";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "ERROR in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            if (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue.ToLower().Contains(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].ExpectedValue.ToLower()) == false)
                            {
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "FAIL";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "FAIL in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            break;
                            #endregion

                        case "ListDir":
                            #region ListDir
                            s_InfoStrUSBStorage = string.Format("** [Step Index]: {0}   [Function Name]:FTP    [Name]:{1} ", st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Index, st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                            sDestination = st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FtpServerFolder;
                            response = false;
                            string sGetDir = "";
                            try
                            {
                                response = ftpclient.ListDirUSBStorage(sDestination, ref sGetDir, ref exceptionInfo);
                                if (response)
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "ListDir successfully";
                                    Invoke(new SetTextCallBack(SetText), new object[] { sGetDir + "\nListDir successfully", txtUSBStorageFunctionTestInformation });
                                }
                                else
                                {
                                    st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = "ListDir fail: " + exceptionInfo;
                                    Invoke(new SetTextCallBack(SetText), new object[] { "ListDir fail: " + exceptionInfo, txtUSBStorageFunctionTestInformation });
                                }
                            }
                            catch (Exception ex)
                            {
                                s_InfoStrUSBStorage = string.Format("**The step error in: {0}", ex);
                                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                                st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue = s_InfoStrUSBStorage;
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "ERROR";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "ERROR in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            if (st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].GetValue.ToLower().Contains(st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].ExpectedValue.ToLower()) == false)
                            {
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "FAIL";
                                st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "FAIL in: \n" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name;
                                foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                                {
                                    cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                                }
                                i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                                break;
                            }
                            break;
                            #endregion
                    }
                    //s_InfoStrUSBStorage = string.Format("** The step was run successfully ");
                    //Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrUSBStorage, txtUSBStorageFunctionTestInformation });
                    break;
                    #endregion

                default:
                    st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].TestResult = "ERROR";
                    st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Comment = "ERROR: \nThere is no function named \"" + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].FunctionName + "\" in the program";
                    foreach (CbtLineNotificationApi cln_LineNotificationUSBStorage in cln_LineNotificationListUSBStorage)
                    {
                        cln_LineNotificationUSBStorage.postMessage("Name: " + st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].Name + ", FAIL in: " + st_ReadStepsScriptDataUSBStorage[i_StepsScriptIndexUSBStorage].Name);
                    }
                    i_duringStartStopUSBStorage = Convert.ToInt32(st_ReadFinalScriptDataUSBStorage[i_FinalScriptIndexUSBStorage].StopIndex) + 1;
                    break;
            }

            bCommonSingleScriptItemRunning = false;
            Thread.Sleep(2000);


        }

        private void ToggleFunctionTestControllerUSBStorage(bool Toggle)
        {
            btnUSBStorageFunctionTestRun.Text = Toggle ? "Run" : "Stop";

            //----- Function Test -----//
            gbox_USBStorageModel.Enabled = Toggle;
            gbox_USBStorageGetewaySetting.Enabled = Toggle;
            gbox_USBStorageFunctionTestScript.Enabled = Toggle;


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


            if (bCommonFTThreadRunning == false)
            {
                btnUSBStorageFunctionTestRun.Enabled = false;
                WriteconsoleLogUSBStorage();
                WriteConsoleLog.Close();
                WriteDebugMsg.Close();
                try
                {
                    thread_WebGUIScriptItem.Abort();
                }
                catch { }
                try
                {
                    threadCommonFTscriptItem.Abort();
                }
                catch { }
                try
                {
                    threadCommonFT.Abort();
                }
                catch { }

                try
                {
                    SetExcelCellsHighUSBStorage();
                    SetWebGUIExcelCellsHighUSBStorage();
                    TestFinishedActionUSBStorage();

                }
                catch { }

                SaveEndTime_and_IEversion_to_Excel(cs_BrowserUSBStorage);
                Save_and_Close_Excel_File();
                File.SetAttributes(txtUSBStorageFunctionTestFinalScriptFilePath.Text, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
                File.SetAttributes(txtUSBStorageFunctionTestStepsScriptFilePath.Text, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
                File.SetAttributes(txtUSBStorageFunctionTestScriptFilePath.Text, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀

                try
                {
                    cs_BrowserUSBStorage.Close_WebDriver();
                }
                catch { }
                //Close_WebDriver_and_Browser();

                //MessageBox.Show(this, "Test complete !!!", "Information", MessageBoxButtons.OK);
                MessageBox.Show(new Form { TopMost = true }, "Test complete !!!", "ATE Information", MessageBoxButtons.OK); // 將 MessageBox於桌面置頂，否則會被 Browser蓋住
                btnUSBStorageFunctionTestRun.Enabled = true;
                bCommonTestComplete = true;
            }
        }

        private void WriteconsoleLogUSBStorage()
        {
            if (ckboxUSBStorageFunctionTestConsoleLog.Checked == true)
            {
                s_ConsoleTextUSBStorage = string.Empty;

                for (int i = 0; i < 2; i++)
                {
                    s_ConsoleTextUSBStorage += ReadConsoleLogByBytesUSBStorage(cp_ComPortUSBStorage);
                    Thread.Sleep(500);
                }

                s_ConsoleTextUSBStorage += "\r\n\r\n\r\n****************************************************************************************************************\r\n\r\n\r\n";
                WriteConsoleLog.Write(s_ConsoleTextUSBStorage);
            }
        }

        private void SetExcelCellsHighUSBStorage()
        {
            string CELL1;
            string CELL2;

            //excelWorkSheetCommon = excelWorkBookCommon.Sheets[2];
            //CELL1 = "B8";
            //CELL2 = "B25";
            //excelRangeCommon = excelWorkSheetCommon.Range[CELL1, CELL2];
            //excelRangeCommon.RowHeight = 15;

            st_ExcelObjectUSBStorage.excelWorkSheet = st_ExcelObjectUSBStorage.excelWorkBook.Sheets[1];
            CELL1 = "B16";
            CELL2 = "B50000";
            st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Range[CELL1, CELL2];
            st_ExcelObjectUSBStorage.excelRange.RowHeight = 15;

        }

        private void SetWebGUIExcelCellsHighUSBStorage()
        {
            string CELL1;
            string CELL2;

            //excelWorkSheetCommon = excelWorkBookCommon.Sheets[2];
            //CELL1 = "B8";
            //CELL2 = "B25";
            //excelRangeCommon = excelWorkSheetCommon.Range[CELL1, CELL2];
            //excelRangeCommon.RowHeight = 15;

            excelWorkSheetCommon = excelWorkBookCommon.Sheets[1];
            CELL1 = "B15";
            CELL2 = "B50000";
            excelRangeCommon = excelWorkSheetCommon.Range[CELL1, CELL2];
            excelRangeCommon.RowHeight = 15;

        }

        private void ToggleFunctionTestUSBStorage()
        {
            ToggleFunctionTestControllerUSBStorage(true);
            Debug.WriteLine("Toggle");
        }

        private void SaveModelParameterUSBStorage()
        {
            mgs_ModelParameterUSBStorage.ModelName = txtUSBStorageFunctionTestModelName.Text;
            mgs_ModelParameterUSBStorage.SN = txtUSBStorageFunctionTestModelSerialNumber.Text;
            mgs_ModelParameterUSBStorage.SWver = txtUSBStorageFunctionTestModelSwVersion.Text;
            mgs_ModelParameterUSBStorage.HWver = txtUSBStorageFunctionTestModelHwVersion.Text;
        }

        private void SaveGatewayLoginParameterUSBStorage()
        {
            //lsgsGatewayLoginSettingParameter.GatewayIP = masktxtUSBStorageFunctionTestGatewayIP.Text;

            ls_LoginSettingParametersUSBStorage.GatewayIP = masktxtUSBStorageFunctionTestGatewayIP.Text;
            ls_LoginSettingParametersUSBStorage.HTTP_Port = nudUSBStorageFunctionTestDefaultSettingsHTTPport.Value.ToString();
            //lsLoginSettingParameters.RebootWaitTime = Convert.ToInt32(nudUSBStorageFunctionTestRebootWaitTime.Value.ToString());

            if (ckboxUSBStorageFunctionTestConsoleLog.Checked)
                ls_LoginSettingParametersUSBStorage.ConsoleLog = true;
            else
                ls_LoginSettingParametersUSBStorage.ConsoleLog = false;
        }

        private void SaveTestBrowserParameterUSBStorage()
        {
            tbs_TestBrowerParametersUSBStorage.TestRun = nudUSBStorageFunctionTestTestRun.Value.ToString();
            if (ckboxUSBStorageFunctionTestStopWhenTestError.Checked)
                tbs_TestBrowerParametersUSBStorage.StopWhenTestError = true;

        }

        private bool CheckTestBrowserUSBStorage()
        {
            csbt_TestBrowserListUSBStorage.Clear();

            /*if (ckbox_USBStorageTestBrowser_Chrome.Checked == false &&
                ckbox_USBStorageFT_TestBrowser_FireFox.Checked == false &&
                ckbox_USBStorageFT_TestBrowser_IE.Checked == false)
            {
                return false;
            }*/

            //if (ckbox_CGA2121FT_TestBrowser_Chrome.Checked == true)
            csbt_TestBrowserListUSBStorage.Add(CBT_SeleniumApi.BrowserType.Chrome);
            //if (ckbox_CGA2121FT_TestBrowser_FireFox.Checked == true)
            //    csbtTestBrowserList.Add(CBT_SeleniumApi.BrowserType.FireFox);
            //if (ckbox_USBStorageFT_TestBrowser_IE.Checked == true)
            //    csbtTestBrowserList.Add(CBT_SeleniumApi.BrowserType.IE);

            return true;
        }

        private void CreateNewExcelFileUSBStorage()
        {
            s_ReportFileNameUSBStorage = "MODEL_FinalReport_SN_SW_HW_DATE.xlsx";
            s_ReportFileNameUSBStorage = getReportFileNameUSBStorage(s_ReportFileNameUSBStorage);

            string sSubPath;
            bool bIsExists;

            s_FolderNameUSBStorage = s_FolderNameUSBStorage.Replace("DATE", DateTime.Now.ToString("yyyy-MMdd-HHmmss"));
            sSubPath = System.Windows.Forms.Application.StartupPath + "\\report\\" + s_FolderNameUSBStorage;
            bIsExists = System.IO.Directory.Exists(sSubPath);

            if (!bIsExists)
                System.IO.Directory.CreateDirectory(sSubPath);

            s_ReportFilePathUSBStorage = s_ReportPathUSBStorage + @"\" + s_FolderNameUSBStorage + @"\" + s_ReportFileNameUSBStorage;
            initialExcel_CommonCaseUSBStorage(s_ReportFilePathUSBStorage);
        }

        private void CreateNewWebGUIExcelFileUSBStorage()
        {
            s_ReportFileNameWebGUIUSBStorage = "MODEL_WebGUIReport_SN_SW_HW_DATE_BROWSER.xlsx";
            s_ReportFileNameWebGUIUSBStorage = getReportFileNameUSBStorage(s_ReportFileNameWebGUIUSBStorage);

            s_ReportFilePathWebGUIUSBStorage = s_ReportPathUSBStorage + @"\" + s_FolderNameUSBStorage + @"\" + s_ReportFileNameWebGUIUSBStorage;
            initialWebGUIExcel_CommonCaseUSBStorage(s_ReportFilePathWebGUIUSBStorage);
        }

        public string getReportFileNameUSBStorage(string reportTargetfileName)
        {
            reportTargetfileName = reportTargetfileName.Replace("MODEL", (mgs_ModelParameterUSBStorage.ModelName == "") ? "USB_Storage" : txtUSBStorageFunctionTestModelName.Text);
            reportTargetfileName = reportTargetfileName.Replace("SN", mgs_ModelParameterUSBStorage.SN);
            reportTargetfileName = reportTargetfileName.Replace("SW", mgs_ModelParameterUSBStorage.SWver);
            reportTargetfileName = reportTargetfileName.Replace("HW", mgs_ModelParameterUSBStorage.HWver);
            reportTargetfileName = reportTargetfileName.Replace("DATE", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            reportTargetfileName = reportTargetfileName.Replace("BROWSER", Convert.ToString(csbt_TestBrowserUSBStorage));
            return reportTargetfileName;
        }

        private void initialExcel_CommonCaseUSBStorage(string filePath)
        {
            st_ExcelObjectUSBStorage.excelApp = new Excel.Application();

            /*** Set Excel visible ***/
            st_ExcelObjectUSBStorage.excelApp.Visible = true;

            /*** Do not show alert ***/
            st_ExcelObjectUSBStorage.excelApp.DisplayAlerts = false;

            st_ExcelObjectUSBStorage.excelApp.UserControl = true;
            st_ExcelObjectUSBStorage.excelApp.Interactive = false;

            //Set font and font size attributes
            st_ExcelObjectUSBStorage.excelApp.StandardFont = "Times New Roman";
            st_ExcelObjectUSBStorage.excelApp.StandardFontSize = 11;

            /*** This method is used to open an Excel workbook by passing the file path as a parameter to this method. ***/
            st_ExcelObjectUSBStorage.excelWorkBook = st_ExcelObjectUSBStorage.excelApp.Workbooks.Add(misValue);

            st_ExcelObjectUSBStorage.excelWorkSheet = st_ExcelObjectUSBStorage.excelWorkBook.Sheets[1];
            st_ExcelObjectUSBStorage.excelWorkSheet.Name = "Test Results";
            createExcelTitleUSBStorage();

            saveExcelUSBStorage(filePath);
        }

        private void initialWebGUIExcel_CommonCaseUSBStorage(string filePath)
        {
            excelAppCommon = new Excel.Application();

            /*** Set Excel visible ***/
            excelAppCommon.Visible = true;

            /*** Do not show alert ***/
            excelAppCommon.DisplayAlerts = false;

            excelAppCommon.UserControl = true;
            excelAppCommon.Interactive = false;

            //Set font and font size attributes
            excelAppCommon.StandardFont = "Times New Roman";
            excelAppCommon.StandardFontSize = 11;

            /*** This method is used to open an Excel workbook by passing the file path as a parameter to this method. ***/
            excelWorkBookCommon = excelAppCommon.Workbooks.Add(misValue);

            excelWorkSheetCommon = excelWorkBookCommon.Sheets[1];
            excelWorkSheetCommon.Name = "Web GUI Test Results";
            createWebGUIExcelTitleUSBStorage();

            saveWebGUIExcelUSBStorage(filePath);
        }

        private void createExcelTitleUSBStorage()
        {
            //*** SettingExcelFont() ***//
            //--- [ColorIndex] 黑:1, 紅:3, 藍:5,淡黃:27, 淡橘:44 
            //--- [Font Style] U:底線 B:粗體 I:斜體 (可同時設定多個屬性)   ex:"UBI"、"BI"、"I"、"B"...

            //st_ExcelObjectUSBStorage.excelApp.Cells[2, 2] = "CyberATE EA8500 USB Storage Test Report";
            st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 2];
            SettingExcelFontUSBStorage(st_ExcelObjectUSBStorage.excelRange, "Times New Roman", 14, "U", 1);
            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 2] = "CyberATE EA8500 USB Storage Final Report";

            st_ExcelObjectUSBStorage.excelApp.Cells[4, 2] = "Test";
            st_ExcelObjectUSBStorage.excelApp.Cells[5, 2] = "Station";
            st_ExcelObjectUSBStorage.excelApp.Cells[6, 2] = "Start Time";
            st_ExcelObjectUSBStorage.excelApp.Cells[7, 2] = "End Time";
            st_ExcelObjectUSBStorage.excelApp.Cells[8, 2] = "Model";
            st_ExcelObjectUSBStorage.excelApp.Cells[9, 2] = "Serial";
            st_ExcelObjectUSBStorage.excelApp.Cells[10, 2] = "SW Version";
            st_ExcelObjectUSBStorage.excelApp.Cells[11, 2] = "HW Version";

            st_ExcelObjectUSBStorage.excelApp.Cells[4, 3] = "EA8500 USB Storage Test";
            st_ExcelObjectUSBStorage.excelApp.Cells[5, 3] = "CyberTAN ATE-" + txtUSBStorageFunctionTestModelName.Text;
            st_ExcelObjectUSBStorage.excelApp.Cells[6, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            st_ExcelObjectUSBStorage.excelApp.Cells[8, 3] = txtUSBStorageFunctionTestModelName.Text;
            st_ExcelObjectUSBStorage.excelApp.Cells[9, 3] = txtUSBStorageFunctionTestModelSerialNumber.Text;
            st_ExcelObjectUSBStorage.excelApp.Cells[10, 3] = txtUSBStorageFunctionTestModelSwVersion.Text;
            st_ExcelObjectUSBStorage.excelApp.Cells[11, 3] = txtUSBStorageFunctionTestModelHwVersion.Text;

            /* Set cells width */
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 1], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 1]].ColumnWidth = 2;
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 2], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 2]].ColumnWidth = 15; /* column B */
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 3], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 3]].ColumnWidth = 15;
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 4], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 4]].ColumnWidth = 15; /* column B */
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 5], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 5]].ColumnWidth = 15;
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 6], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 6]].ColumnWidth = 15; /* column B */
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 7], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 7]].ColumnWidth = 15;
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 8], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 8]].ColumnWidth = 15; /* column B */
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 9], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 9]].ColumnWidth = 15; /* column B */
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 10], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 10]].ColumnWidth = 15; /* column B */
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[2, 11], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 11]].ColumnWidth = 15; /* column B */

            st_ExcelObjectUSBStorage.excelApp.Cells[3, 8] = "Option";
            st_ExcelObjectUSBStorage.excelApp.Cells[3, 9] = "Value";

            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelApp.Cells[3, 8], st_ExcelObjectUSBStorage.excelApp.Cells[7, 9]].Borders.LineStyle = 1;

            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelApp.Cells[3, 8], st_ExcelObjectUSBStorage.excelApp.Cells[3, 9]].Font.Underline = true;
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelApp.Cells[3, 8], st_ExcelObjectUSBStorage.excelApp.Cells[3, 9]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelApp.Cells[3, 8], st_ExcelObjectUSBStorage.excelApp.Cells[3, 9]].Font.FontStyle = "Bold";

            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[4, 8] = "Loop";
            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[4, 9] = i_TestRun;
            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[5, 8] = "Final Script";
            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[5, 9] = txtUSBStorageFunctionTestFinalScriptFilePath.Text;
            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[6, 8] = "Steps Script";
            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[6, 9] = txtUSBStorageFunctionTestStepsScriptFilePath.Text;
            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[7, 8] = "WebGUI Script";
            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[7, 9] = txtUSBStorageFunctionTestScriptFilePath.Text;

            st_ExcelObjectUSBStorage.excelApp.Cells[15, 2] = "Item No";
            st_ExcelObjectUSBStorage.excelApp.Cells[15, 3] = "Function Name";
            st_ExcelObjectUSBStorage.excelApp.Cells[15, 4] = "Name";
            st_ExcelObjectUSBStorage.excelApp.Cells[15, 5] = "Test Result";
            st_ExcelObjectUSBStorage.excelApp.Cells[15, 6] = "Comment";
            st_ExcelObjectUSBStorage.excelApp.Cells[15, 7] = "Log File";
            //st_ExcelObjectUSBStorage.excelApp.Cells[15, 8] = "";

            st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelWorkSheet.Cells[14, 2], st_ExcelObjectUSBStorage.excelWorkSheet.Cells[14, 10]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


            //st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelApp.Cells[15, 2], st_ExcelObjectUSBStorage.excelApp.Cells[15, 10]].Font.FontStyle = "Bold";
            //st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelApp.Cells[15, 2], st_ExcelObjectUSBStorage.excelApp.Cells[15, 10]];
            //SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "C");


            ////*** SettingExcelAlignment() ***//
            ////--- [direction] H:水平方向 V:垂直方向 (可同時設定多個屬性)   ex:"HV"、"H"、"V"
            ////--- [AlignmentType] C:置中 L:靠左 R:靠右

            ////st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelApp.Cells[6, 3], st_ExcelObjectUSBStorage.excelApp.Cells[11, 3]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Range[st_ExcelObjectUSBStorage.excelApp.Cells[4, 3], st_ExcelObjectUSBStorage.excelApp.Cells[12, 3]];
            //SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "C");

            //st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Range["E4", "E11"];
            //SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "C");


            /*** Set cells width ***/
            //SetExcelCellsWidthUSBStorage();

        }

        private void createWebGUIExcelTitleUSBStorage()
        {
            //*** SettingExcelFont() ***//
            //--- [ColorIndex] 黑:1, 紅:3, 藍:5,淡黃:27, 淡橘:44 
            //--- [Font Style] U:底線 B:粗體 I:斜體 (可同時設定多個屬性)   ex:"UBI"、"BI"、"I"、"B"...

            //excelAppCommon.Cells[2, 2] = "CyberATE EA8500 USB Storage Test Report";
            excelRangeCommon = excelWorkSheetCommon.Cells[2, 2];
            SettingExcelFontUSBStorage(excelRangeCommon, "Times New Roman", 14, "U", 1);
            excelWorkSheetCommon.Cells[2, 2] = "CyberATE EA8500 USB Storage WebGUI Test Report";

            excelAppCommon.Cells[4, 2] = "Test";
            excelAppCommon.Cells[5, 2] = "Station";
            excelAppCommon.Cells[6, 2] = "Start Time";
            excelAppCommon.Cells[7, 2] = "End Time";
            excelAppCommon.Cells[8, 2] = "Model";
            excelAppCommon.Cells[9, 2] = "Serial";
            excelAppCommon.Cells[10, 2] = "SW Version";
            excelAppCommon.Cells[11, 2] = "HW Version";
            //excelAppCommon.Cells[12, 2] = "Test Browser";

            excelAppCommon.Cells[4, 3] = "EA8500 USB Storage Test";
            excelAppCommon.Cells[5, 3] = "CyberTAN ATE-" + txtUSBStorageFunctionTestModelName.Text;
            excelAppCommon.Cells[6, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            excelAppCommon.Cells[8, 3] = txtUSBStorageFunctionTestModelName.Text;
            excelAppCommon.Cells[9, 3] = txtUSBStorageFunctionTestModelSerialNumber.Text;
            excelAppCommon.Cells[10, 3] = txtUSBStorageFunctionTestModelSwVersion.Text;
            excelAppCommon.Cells[11, 3] = txtUSBStorageFunctionTestModelHwVersion.Text;

            excelAppCommon.Cells[14, 2] = "Test Index";
            excelAppCommon.Cells[14, 3] = "Action Name";
            excelAppCommon.Cells[14, 4] = "Action";
            excelAppCommon.Cells[14, 5] = "Element Type";
            excelAppCommon.Cells[14, 6] = "Write Value";
            excelAppCommon.Cells[14, 7] = "Expected Value";
            excelAppCommon.Cells[14, 8] = "Get Value";
            excelAppCommon.Cells[14, 9] = "Rsult";
            excelAppCommon.Cells[14, 10] = "Note";

            excelWorkSheetCommon.Range[excelAppCommon.Cells[14, 2], excelAppCommon.Cells[14, 10]].Font.FontStyle = "Bold";
            excelRangeCommon = excelWorkSheetCommon.Range[excelAppCommon.Cells[14, 2], excelAppCommon.Cells[14, 10]];
            SettingExcelAlignment(excelRangeCommon, "H", "C");


            //*** SettingExcelAlignment() ***//
            //--- [direction] H:水平方向 V:垂直方向 (可同時設定多個屬性)   ex:"HV"、"H"、"V"
            //--- [AlignmentType] C:置中 L:靠左 R:靠右

            //excelWorkSheetCommon.Range[excelAppCommon.Cells[6, 3], excelAppCommon.Cells[11, 3]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            excelRangeCommon = excelWorkSheetCommon.Range[excelAppCommon.Cells[4, 3], excelAppCommon.Cells[12, 3]];
            SettingExcelAlignment(excelRangeCommon, "H", "C");

            excelRangeCommon = excelWorkSheetCommon.Range["E4", "E11"];
            SettingExcelAlignment(excelRangeCommon, "H", "C");


            /*** Set cells width ***/
            SetExcelCellsWidthUSBStorage();

        }

        private void SetExcelCellsWidthUSBStorage()
        {
            excelRangeCommon = excelWorkSheetCommon.Cells[1, 1];
            excelRangeCommon.ColumnWidth = 2;

            excelRangeCommon = excelWorkSheetCommon.Cells[1, 2];
            excelRangeCommon.ColumnWidth = 13;

            excelRangeCommon = excelWorkSheetCommon.Cells[1, 3];
            excelRangeCommon.ColumnWidth = 35;

            excelRangeCommon = excelWorkSheetCommon.Cells[1, 4];
            excelRangeCommon.ColumnWidth = 9;

            excelRangeCommon = excelWorkSheetCommon.Cells[1, 5];
            excelRangeCommon.ColumnWidth = 20;

            excelRangeCommon = excelWorkSheetCommon.Cells[1, 6];
            excelRangeCommon.ColumnWidth = 16;

            excelRangeCommon = excelWorkSheetCommon.Cells[1, 7];
            excelRangeCommon.ColumnWidth = 16;

            excelRangeCommon = excelWorkSheetCommon.Cells[1, 8];
            excelRangeCommon.ColumnWidth = 16;

            excelRangeCommon = excelWorkSheetCommon.Cells[1, 9];
            excelRangeCommon.ColumnWidth = 10;

            excelRangeCommon = excelWorkSheetCommon.Cells[1, 10];
            excelRangeCommon.ColumnWidth = 26;

            // 表格框線
            //excelWorkSheetCommon.Range[excelAppCommon.Cells[3, 8], excelAppCommon.Cells[9, 9]].Borders.LineStyle = 1;

            // 底線
            //excelWorkSheetCommon.Range[excelAppCommon.Cells[3, 8], excelAppCommon.Cells[3, 9]].Font.Underline = true;

            // 粗體
            //excelWorkSheetCommon.Range[excelAppCommon.Cells[3, 8], excelAppCommon.Cells[3, 9]].Font.FontStyle = "Bold";

            // 水平置中
            //excelWorkSheetCommon.Range[excelAppCommon.Cells[3, 8], excelAppCommon.Cells[3, 9]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

        }

        private void SettingExcelFontUSBStorage(Excel.Range excelRangeCommon, string FONT_TYPE, int FONT_SIZE, string FONT_STYLE, int COLOR_INDEX)
        {
            //--- [ColorIndex] 黑:1, 紅:3, 藍:5,淡黃:27, 淡橘:44 
            //--- [Font Style] U:底線 B:粗體 I:斜體 (可同時設定多個屬性)   ex:"UBI"、"BI"、"I"、"B"...

            excelRangeCommon.Font.Name = FONT_TYPE;
            excelRangeCommon.Font.Size = FONT_SIZE;
            excelRangeCommon.Font.ColorIndex = COLOR_INDEX;

            if (FONT_STYLE.ToLower().IndexOf("b") >= 0)
                excelRangeCommon.Font.Bold = true;
            if (FONT_STYLE.ToLower().IndexOf("i") >= 0)
                excelRangeCommon.Font.Italic = true;
            if (FONT_STYLE.ToLower().IndexOf("u") >= 0)
                excelRangeCommon.Font.Underline = true;
        }

        private void WriteLogTitleUSBStorage(System.IO.StreamWriter WriteLog, string titleStr, string symbol)
        {
            WriteLog.WriteLine("");
            WriteLog.WriteLine("");

            for (int i = 0; i < 118; i++)
                WriteLog.Write(symbol);

            WriteLog.WriteLine("");
            WriteLog.WriteLine(titleStr);

            for (int i = 0; i < 118; i++)
                WriteLog.Write(symbol);

            WriteLog.WriteLine("");
            WriteLog.WriteLine("");
            WriteLog.WriteLine("");
        }

        private void TestInitialUSBStorage()
        {
            cs_BrowserUSBStorage = new CBT_SeleniumApi();

            if (!cs_BrowserUSBStorage.init(csbt_TestBrowserUSBStorage))
            {
                SetExcelCellsHighUSBStorage();
                SetWebGUIExcelCellsHighUSBStorage();
                SaveEndTime_and_Version_to_ExcelUSBStorage();
                SaveEndTime_and_Version_to_WebGUIExcelUSBStorage();
                Save_and_Close_Excel_File();
                Save_and_Close_Excel_File_Common(st_ExcelObjectUSBStorage);
                cs_BrowserUSBStorage.Close_WebDriver();
                bCommonFTThreadRunning = false;
                bWebGUISingleScriptItemRunning = false;
                bCommonSingleScriptItemRunning = false;
                return;
            }
            cs_BrowserUSBStorage.SettingTimeout(60);
            cs_BrowserUSBStorage.WindowMaximize();
            //Thread.Sleep(1000);
            //cs_BrowserUSBStorage.WindowMinimize();
            Thread.Sleep(3000);
        }

        private void SaveEndTime_and_Version_to_ExcelUSBStorage()
        {
            try
            {
                st_ExcelObjectUSBStorage.excelWorkSheet = st_ExcelObjectUSBStorage.excelWorkBook.Sheets[1];

                //st_ExcelObjectUSBStorage.excelApp.Cells[10, 4] = "Script Path";
                //st_ExcelObjectUSBStorage.excelApp.Cells[10, 5] = txtUSBStorageFunctionTestFinalScriptFilePath.Text;
                //st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Cells[10, 5];
                //SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "R");
                //st_ExcelObjectUSBStorage.excelApp.Cells[11, 4] = "Browser Version";
                //st_ExcelObjectUSBStorage.excelApp.Cells[11, 5] = cs_BrowserUSBStorage.BrowserVersion();
                //st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Cells[11, 5];
                //SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "R");

                st_ExcelObjectUSBStorage.excelApp.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");  // Save End Time to Excel
            }
            catch { }

        }

        private void SaveEndTime_and_Version_to_WebGUIExcelUSBStorage()
        {
            try
            {
                excelWorkSheetCommon = excelWorkBookCommon.Sheets[1];

                excelAppCommon.Cells[10, 4] = "Script Path";
                excelAppCommon.Cells[10, 5] = txtUSBStorageFunctionTestScriptFilePath.Text;
                excelRangeCommon = excelWorkSheetCommon.Cells[10, 5];
                SettingExcelAlignment(excelRangeCommon, "H", "R");
                excelAppCommon.Cells[11, 4] = "Browser Version";

                excelAppCommon.Cells[11, 5] = cs_BrowserUSBStorage.BrowserVersion();

                excelRangeCommon = excelWorkSheetCommon.Cells[11, 5];
                SettingExcelAlignment(excelRangeCommon, "H", "R");



                excelAppCommon.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");  // Save End Time to Excel
            }
            catch { }

        }

        private bool CheckExtraTestStatusUSBStorage(ref string TestStatus)
        {
            //-----------------------------------------------//
            //-------------- Check Ping Status --------------//
            //-----------------------------------------------//
            bool pingStatus = false;
            Stopwatch pingTimer = new Stopwatch();
            pingTimer.Reset();
            pingTimer.Start();
            Invoke(new SetTextCallBack(SetText), new object[] { "\tCheck Ping Status", txtUSBStorageFunctionTestInformation });
            while (true)
            {
                if (pingTimer.ElapsedMilliseconds > (Convert.ToInt32(5000)))
                {
                    break;
                }

                if (PingClient(ls_LoginSettingParametersUSBStorage.GatewayIP, 1000))
                {
                    pingStatus = true;
                    TestStatus = "Ping Successfully";
                    Invoke(new SetTextCallBack(SetText), new object[] { "\tPing Successfully!!", txtUSBStorageFunctionTestInformation });
                    break;
                }

                Thread.Sleep(1000);
                string Info = String.Format(".");
                Invoke(new SetTextCallBack(SetText), new object[] { Info, txtUSBStorageFunctionTestInformation });
            }

            if (!pingStatus)
            {
                TestStatus = "Ping Failed";
                Invoke(new SetTextCallBack(SetText), new object[] { "\tPing Failed!!", txtUSBStorageFunctionTestInformation });
                return false;
            }





            return true;
        }

        private void WriteTestReportUSBStorage(bool TestResult, FinalScriptDataUSBStorage FinalScriptPara)
        {

            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 2] = FinalScriptPara.TestItem;
            st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 2];
            SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "C");

            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 3] = FinalScriptPara.FunctionName;
            st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 3];
            SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "L");


            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 4] = FinalScriptPara.Name;
            st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 4];
            SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "L");

            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 5] = FinalScriptPara.TestResult;
            st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 5];
            SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "C");
            if (FinalScriptPara.TestResult == "PASS")
            {
                st_ExcelObjectUSBStorage.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
            }
            else if (FinalScriptPara.TestResult == "FAIL" || FinalScriptPara.TestResult == "ERROR")
            {
                st_ExcelObjectUSBStorage.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }


            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 6] = FinalScriptPara.Comment;
            st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 6];
            SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "C");


            st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 7] = FinalScriptPara.Log;
            st_ExcelObjectUSBStorage.excelRange = st_ExcelObjectUSBStorage.excelWorkSheet.Cells[i_DATA_FinalROWUSBStorage, 7];
            //st_ExcelObjectUSBStorage.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            SettingExcelAlignment(st_ExcelObjectUSBStorage.excelRange, "H", "C");
            st_ExcelObjectUSBStorage.excelWorkSheet.Hyperlinks.Add(st_ExcelObjectUSBStorage.excelRange, s_LogFileUSBStorage, Type.Missing, "ConsoleLog", "ConsoleLog");

            SettingExcelFontUSBStorage(st_ExcelObjectUSBStorage.excelRange, "Times New Roman", 11, "B", 3);
            i_DATA_FinalROWUSBStorage++;
        }

        private void WriteWebGUITestReportUSBStorage(int Index, bool TestResult, string ExceptionInfo)
        {
            string Action = st_ReadScriptDataUSBStorage[Index].Action;

            excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 2] = st_ReadScriptDataUSBStorage[Index].TestIndex;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 2];
            SettingExcelAlignment(excelRangeCommon, "H", "C");

            excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 3] = st_ReadScriptDataUSBStorage[Index].ActionName;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 3];
            SettingExcelAlignment(excelRangeCommon, "H", "L");

            excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 4] = Action;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 4];
            SettingExcelAlignment(excelRangeCommon, "H", "C");

            excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 5] = st_ReadScriptDataUSBStorage[Index].ElementType;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 5];
            SettingExcelAlignment(excelRangeCommon, "H", "L");

            excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 10] = ExceptionInfo;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 10];
            SettingExcelAlignment(excelRangeCommon, "H", "L");
            if (ExceptionInfo.CompareTo("Time Out") == 0)
                SettingExcelFontUSBStorage(excelRangeCommon, "Times New Roman", 11, "B", 3);


            if (Action.CompareTo("Set") == 0)
            {
                excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 6] = st_ReadScriptDataUSBStorage[Index].WriteExpectedValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 6];
                SettingExcelAlignment(excelRangeCommon, "H", "C");
            }
            else if (Action.CompareTo("Get") == 0)
            {
                excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 7] = st_ReadScriptDataUSBStorage[Index].WriteExpectedValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 7];
                SettingExcelAlignment(excelRangeCommon, "H", "C");

                excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 8] = st_ReadScriptDataUSBStorage[Index].GetValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 8];
                SettingExcelAlignment(excelRangeCommon, "H", "C");

                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 9];
                if (TestResult == true)
                {
                    excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 9] = "Pass";
                    SettingExcelFontUSBStorage(excelRangeCommon, "Times New Roman", 11, "B", 5);
                }
                else
                {
                    excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 9] = "Fail";
                    SettingExcelFontUSBStorage(excelRangeCommon, "Times New Roman", 11, "B", 3);
                }
                SettingExcelAlignment(excelRangeCommon, "H", "C");
            }

            i_DATA_ROWUSBStorage = i_DATA_ROWUSBStorage + 1;
        }

        private void WriteWebGUITestReportUSBStorage(bool TestResult, CBT_SeleniumApi.GuiScriptParameter ScriptPara)
        {
            string Action = ScriptPara.Action;

            excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 2] = ScriptPara.Index;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 2];
            SettingExcelAlignment(excelRangeCommon, "H", "C");

            excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 3] = ScriptPara.ActionName;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 3];
            SettingExcelAlignment(excelRangeCommon, "H", "L");

            excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 4] = Action;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 4];
            SettingExcelAlignment(excelRangeCommon, "H", "C");

            excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 5] = ScriptPara.ElementType;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 5];
            SettingExcelAlignment(excelRangeCommon, "H", "L");

            excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 10] = ScriptPara.Note;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 10];
            SettingExcelAlignment(excelRangeCommon, "H", "L");
            if (ScriptPara.Note.CompareTo("Time Out") == 0)
                SettingExcelFontUSBStorage(excelRangeCommon, "Times New Roman", 11, "B", 3);


            if (Action.CompareTo("Set") == 0)
            {
                excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 6] = ScriptPara.WriteValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 6];
                SettingExcelAlignment(excelRangeCommon, "H", "C");
            }
            else if (Action.CompareTo("Get") == 0)
            {
                excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 7] = ScriptPara.ExpectedValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 7];
                SettingExcelAlignment(excelRangeCommon, "H", "C");

                excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 8] = ScriptPara.GetValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 8];
                SettingExcelAlignment(excelRangeCommon, "H", "C");

                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 9];
                if (TestResult == true)
                {
                    excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 9] = "Pass";
                    SettingExcelFontUSBStorage(excelRangeCommon, "Times New Roman", 11, "B", 5);
                }
                else
                {
                    excelWorkSheetCommon.Cells[i_DATA_ROWUSBStorage, 9] = "Fail";
                    SettingExcelFontUSBStorage(excelRangeCommon, "Times New Roman", 11, "B", 3);
                }
                SettingExcelAlignment(excelRangeCommon, "H", "C");
            }

            i_DATA_ROWUSBStorage = i_DATA_ROWUSBStorage + 1;
        }

        private bool StopWhenTestErrorUSBStorage()
        {
            if (tbs_TestBrowerParametersUSBStorage.StopWhenTestError == true)
                return true;
            return false;
        }

        private void TestFinishedActionUSBStorage()
        {
            SetExcelCellsHighUSBStorage();
            SaveEndTime_and_Version_to_ExcelUSBStorage();
            Save_and_Close_Excel_File_Common(st_ExcelObjectUSBStorage);
            Thread.Sleep(3000);
        }

        private void TestFinishedWebGUIActionUSBStorage()
        {
            SetWebGUIExcelCellsHighUSBStorage();
            SaveEndTime_and_Version_to_WebGUIExcelUSBStorage();
            Save_and_Close_Excel_File();

            try
            {
                cs_BrowserUSBStorage.Close_WebDriver();
            }
            catch { }
            Thread.Sleep(3000);
        }

        private bool Re_LoginUSBStorage()
        {
            string sCurrent_URL = string.Empty;
            string sElementType = string.Empty;
            string sElementXpath = string.Empty;
            string sElementXpathforCheck = string.Empty;
            string sWriteValue = string.Empty;
            string sExceptionInfo = string.Empty;
            CBT_SeleniumApi.GuiScriptParameter ScriptPara = new CBT_SeleniumApi.GuiScriptParameter();

            try
            {

            }
            catch
            {
                return false;
            }

            return true;
        }

        private void SwitchToRGconsoleUSBStorage()
        {
            if (ckboxUSBStorageFunctionTestConsoleLog.Checked == true)
            {
                string ConText = string.Empty;
                string newLine = "\r\n\r\n\r\n";

                cp_ComPortUSBStorage.Write(newLine);
                Thread.Sleep(500);
                ConText += ReadConsoleLogByBytesUSBStorage(cp_ComPortUSBStorage);
                //MessageBox.Show("ConText1:"+ ConText);
                if (ConText.ToLower().IndexOf("cm>") >= 0)
                {
                    cp_ComPortUSBStorage.Write("cd Con\r\n");
                    Thread.Sleep(500);
                }

                ConText = string.Empty;
                cp_ComPortUSBStorage.Write(newLine);
                Thread.Sleep(500);
                ConText += ReadConsoleLogByBytesUSBStorage(cp_ComPortUSBStorage);
                //MessageBox.Show("ConText2:"+ ConText);
                if (ConText.ToLower().IndexOf("cm/cm_console>") >= 0)
                {
                    cp_ComPortUSBStorage.Write("switch\r\n");
                    Thread.Sleep(500);
                }

                ConText = string.Empty;
                cp_ComPortUSBStorage.Write(newLine);
                Thread.Sleep(500);
                ConText += ReadConsoleLogByBytesUSBStorage(cp_ComPortUSBStorage);
                //MessageBox.Show("ConText3:"+ConText);
                if (ConText.ToLower().IndexOf("rg>") >= 0)
                {
                    cp_ComPortUSBStorage.Write("cd Con\r\n");
                    Thread.Sleep(500);
                }
            }
        }

        private bool CheckLastDataUSBStorage(int DataIndex)
        {
            try
            {
                string NextDataTestIndex = st_ReadScriptDataUSBStorage[DataIndex + 1].TestIndex;
                if (st_ReadScriptDataUSBStorage[DataIndex].TestIndex.CompareTo(NextDataTestIndex) != 0)
                    return true;
            }
            catch
            {
                int LastDataIndex = st_ReadScriptDataUSBStorage.Length - 1;
                if (DataIndex == LastDataIndex)
                    return true;
            }
            return false;
        }

        private bool ExceptionActionUSBStorage(string PageURL, ref CBT_SeleniumApi.GuiScriptParameter ScriptPara)
        {
            if (Re_LoginUSBStorage())
            {
                try
                {
                    cs_BrowserUSBStorage.GoToURL(PageURL);
                }
                catch (Exception ex)
                {
                    ScriptPara.Note = string.Format("ExceptionAction() Error:\nGo To URL {0} Error: \n{1}", PageURL, ex.ToString());
                    return false;
                }
            }
            else
            {
                ScriptPara.Note = "Login Error!!";
                return false;
            }
            return true;
        }

        private void loadFinalScript()
        {
            string filePath = txtUSBStorageFunctionTestFinalScriptFilePath.Text;
            File.SetAttributes(filePath, FileAttributes.ReadOnly); // 設定成唯讀
            excelReadAppCommon = new Excel.Application();
            excelReadWorkBookCommon = excelReadAppCommon.Workbooks.Open(filePath);              //開啟舊檔案
            excelReadWorkSheetCommon = excelReadWorkBookCommon.Sheets[1];
            excelReadRangeCommon = excelReadWorkSheetCommon.UsedRange;
            //File.SetAttributes(filePath, FileAttributes.ReadOnly); // 先開啟檔案再設定成唯讀，可讓有連結的表格更新

            int ColumnCount = excelReadRangeCommon.Columns.Count;
            int RowCount = excelReadRangeCommon.Rows.Count;
            int ActualDataCount = 0;
            string[,] ScriptDataArray = new string[RowCount, ColumnCount];
            int rowOffset = 2;
            int colOffset = 1;
            for (int itemRow = 2; itemRow <= excelReadRangeCommon.Rows.Count; itemRow++)
            {
                //如果第一欄沒打勾或任何標記將不讀取
                if (excelReadRangeCommon[1][itemRow].Text.ToLower() != ("v") || excelReadRangeCommon[1][itemRow].Text == "" || excelReadRangeCommon[1][itemRow].Text == null)
                {
                    rowOffset++;
                    continue;
                }
                for (int itemColumn = 1; itemColumn < excelReadRangeCommon.Columns.Count; itemColumn++)
                {

                    int R_INDEX = itemRow - rowOffset;
                    int C_INDEX = itemColumn - colOffset;
                    ScriptDataArray[R_INDEX, C_INDEX] = excelReadRangeCommon[itemColumn + 1][itemRow].Text;
                }
            }
            closeExcelReadFileCommon();

            ActualDataCount = RowCount - rowOffset + 1;
            st_ReadFinalScriptDataUSBStorage = new FinalScriptDataUSBStorage[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadFinalScriptDataUSBStorage[SCRIPT_INDEX].TestItem = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadFinalScriptDataUSBStorage[SCRIPT_INDEX].FunctionName = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadFinalScriptDataUSBStorage[SCRIPT_INDEX].Name = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadFinalScriptDataUSBStorage[SCRIPT_INDEX].StartIndex = ScriptDataArray[SCRIPT_INDEX, 3];
                st_ReadFinalScriptDataUSBStorage[SCRIPT_INDEX].StopIndex = ScriptDataArray[SCRIPT_INDEX, 4];
                st_ReadFinalScriptDataUSBStorage[SCRIPT_INDEX].TestResult = "";
                st_ReadFinalScriptDataUSBStorage[SCRIPT_INDEX].Comment = "";
                st_ReadFinalScriptDataUSBStorage[SCRIPT_INDEX].Log = "";
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
        }

        private void loadStepsScript()
        {
            string filePath = txtUSBStorageFunctionTestStepsScriptFilePath.Text;
            File.SetAttributes(filePath, FileAttributes.ReadOnly); // 屬性設定成唯讀
            excelReadAppCommon = new Excel.Application();
            excelReadWorkBookCommon = excelReadAppCommon.Workbooks.Open(filePath);              //開啟舊檔案
            excelReadWorkSheetCommon = excelReadWorkBookCommon.Sheets[1];
            excelReadRangeCommon = excelReadWorkSheetCommon.UsedRange;

            int ColumnCount = excelReadRangeCommon.Columns.Count;
            int RowCount = excelReadRangeCommon.Rows.Count;
            int ActualDataCount = 0;
            string[,] ScriptDataArray = new string[RowCount, ColumnCount];
            int rowOffset = 2;
            int colOffset = 1;
            int indexOutNumber = 0;

            for (int itemRow = 2; itemRow <= excelReadRangeCommon.Rows.Count; itemRow++)
            {
                //以下 看到第一欄不是數字就跳下一列，並且這列不算
                if (excelReadRangeCommon[1][itemRow].Text == "" || excelReadRangeCommon[1][itemRow].Text == null || Int32.TryParse(excelReadRangeCommon[1][itemRow].Text, out indexOutNumber) == false)
                {
                    rowOffset++;
                    continue;
                }
                for (int itemColumn = 1; itemColumn <= excelReadRangeCommon.Columns.Count; itemColumn++)
                {
                    int R_INDEX = itemRow - rowOffset;
                    int C_INDEX = itemColumn - colOffset;
                    ScriptDataArray[R_INDEX, C_INDEX] = excelReadRangeCommon[itemColumn][itemRow].Text;
                }
            }
            closeExcelReadFileCommon();

            ActualDataCount = RowCount - rowOffset + 1;
            st_ReadStepsScriptDataUSBStorage = new StepsScriptDataUSBStorage[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].Index = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].FunctionName = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].Name = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].Command = new string[20];
                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].WriteValue = new string[20];
                int writeValuelistIndex = 0;
                int commandlistIndex = 0;
                for (int DC = 3; DC < ColumnCount; DC++)
                {
                    if (ScriptDataArray[SCRIPT_INDEX, DC] != null && ScriptDataArray[SCRIPT_INDEX, DC] != "")
                    {
                        string tmpData = ScriptDataArray[SCRIPT_INDEX, DC].Replace("::", "$");
                        switch (tmpData.Split('$')[0].Trim())
                        {
                            case "Action":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].Action = tmpData.Split('$')[1].Trim();
                                break;

                            case "Command":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].Command[commandlistIndex] = tmpData.Split('$')[1].Trim();
                                commandlistIndex++;
                                break;

                            case "WriteValue":
                                //string[] stmpData = tmpData.Split('$');
                                //int tmpDataLength = stmpData.Length;
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].WriteValue[writeValuelistIndex] = tmpData.Split('$')[tmpData.Split('$').Length - 1].Trim();
                                writeValuelistIndex++;
                                break;

                            case "ExpectedValue":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].ExpectedValue = tmpData.Split('$')[1].Trim();
                                break;

                            case "TimeOut":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].TestTimeOut = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpUsername":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].FtpUsername = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpPassword":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].FtpPassword = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpIP":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].FtpIP = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpServerFolder":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].FtpServerFolder = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpServerPort":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].FtpServerPort = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpRenameTo":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].FtpRenameTo = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpPassiveMode":
                                st_ReadStepsScriptDataUSBStorage[SCRIPT_INDEX].FtpPassiveMode = tmpData.Split('$')[1].Trim();
                                break;
                        }
                    }
                }
            }
            //progressbarForm.Dispose();
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀

        }

        private void loadWebGUIScript()
        {
            string filePath = txtUSBStorageFunctionTestScriptFilePath.Text;
            File.SetAttributes(filePath, FileAttributes.ReadOnly); // 屬性設定成唯讀
            excelReadAppCommon = new Excel.Application();
            excelReadWorkBookCommon = excelReadAppCommon.Workbooks.Open(filePath);              //開啟舊檔案
            excelReadWorkSheetCommon = excelReadWorkBookCommon.Sheets[1];
            excelReadRangeCommon = excelReadWorkSheetCommon.UsedRange;

            int ColumnCount = excelReadRangeCommon.Columns.Count;
            int RowCount = excelReadRangeCommon.Rows.Count;
            int ActualDataCount = 0;
            string[,] ScriptDataArray = new string[RowCount, ColumnCount];
            int rowOffset = 2;
            int colOffset = 1;
            for (int itemRow = 2; itemRow <= excelReadRangeCommon.Rows.Count; itemRow++)
            {
                //如果第一欄沒打勾或任何標記將不讀取  ※excel的行列與Array相反
                if (excelReadRangeCommon[1][itemRow].Text == "" || excelReadRangeCommon[1][itemRow].Text == null)
                {
                    rowOffset++;
                    continue;
                }
                for (int itemColumn = 1; itemColumn <= excelReadRangeCommon.Columns.Count; itemColumn++)
                {

                    int R_INDEX = itemRow - rowOffset;
                    int C_INDEX = itemColumn - colOffset;
                    ScriptDataArray[R_INDEX, C_INDEX] = excelReadRangeCommon[itemColumn][itemRow].Text;
                }
            }
            closeExcelReadFileCommon();

            ActualDataCount = RowCount - rowOffset + 1;
            st_ReadScriptDataUSBStorage = new ScriptDataUSBStorage[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadScriptDataUSBStorage[SCRIPT_INDEX].Procedure = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadScriptDataUSBStorage[SCRIPT_INDEX].TestIndex = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadScriptDataUSBStorage[SCRIPT_INDEX].Action = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadScriptDataUSBStorage[SCRIPT_INDEX].ActionName = ScriptDataArray[SCRIPT_INDEX, 3];
                st_ReadScriptDataUSBStorage[SCRIPT_INDEX].ElementType = ScriptDataArray[SCRIPT_INDEX, 4];
                st_ReadScriptDataUSBStorage[SCRIPT_INDEX].WriteExpectedValue = ScriptDataArray[SCRIPT_INDEX, 5];
                st_ReadScriptDataUSBStorage[SCRIPT_INDEX].RadioButtonExpectedValueXpath = ScriptDataArray[SCRIPT_INDEX, 6];
                st_ReadScriptDataUSBStorage[SCRIPT_INDEX].ElementXpath = ScriptDataArray[SCRIPT_INDEX, 7];
                st_ReadScriptDataUSBStorage[SCRIPT_INDEX].TestTimeOut = ScriptDataArray[SCRIPT_INDEX, 8];
                st_ReadScriptDataUSBStorage[SCRIPT_INDEX].GetValue = "";
            }

            //MessageBox.Show("sdReadScriptData.Length: " + sdReadScriptData.Length);
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
        }

        private void USBStorageFunctionTestLineEnable()
        {
            if (ckboxUSBStorageFunctionTestLineEnable.Checked == true)
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(System.Windows.Forms.Application.StartupPath + @"\config\Line_Notify_USB Storage.xml");

                XmlNode node = doc.SelectSingleNode("CybertanATE");
                if (node == null)
                {
                }

                XmlElement element = (XmlElement)node;
                string strID = element.GetAttribute("Item");
                Debug.Write(strID);
                if (strID.CompareTo("USB Storage") != 0)
                {
                    MessageBox.Show("This XML file is incorrect", "Error");
                }
                XmlNode nodeAttenuator = doc.SelectSingleNode("/CybertanATE/Groups");

                try
                {
                    cln_LineNotificationListUSBStorage.Clear();

                    string ATE_Notify_1 = nodeAttenuator.SelectSingleNode("ATE_Notify_1").InnerText;
                    string ATE_Notify_2 = nodeAttenuator.SelectSingleNode("ATE_Notify_2").InnerText;
                    string ATE_Notify_3 = nodeAttenuator.SelectSingleNode("ATE_Notify_3").InnerText;
                    string ATE_Notify_4 = nodeAttenuator.SelectSingleNode("ATE_Notify_4").InnerText;
                    string ATE_Notify_5 = nodeAttenuator.SelectSingleNode("ATE_Notify_5").InnerText;
                    string ATE_Notify_6 = nodeAttenuator.SelectSingleNode("ATE_Notify_6").InnerText;

                    if (ATE_Notify_1 == "true")
                        cln_LineNotificationListUSBStorage.Add(new CbtLineNotificationApi(CbtLineNotificationApi.dic_LineAuth[1]));

                    if (ATE_Notify_2 == "true")
                        cln_LineNotificationListUSBStorage.Add(new CbtLineNotificationApi(CbtLineNotificationApi.dic_LineAuth[2]));

                    if (ATE_Notify_3 == "true")
                        cln_LineNotificationListUSBStorage.Add(new CbtLineNotificationApi(CbtLineNotificationApi.dic_LineAuth[3]));

                    if (ATE_Notify_4 == "true")
                        cln_LineNotificationListUSBStorage.Add(new CbtLineNotificationApi(CbtLineNotificationApi.dic_LineAuth[4]));

                    if (ATE_Notify_5 == "true")
                        cln_LineNotificationListUSBStorage.Add(new CbtLineNotificationApi(CbtLineNotificationApi.dic_LineAuth[5]));

                    if (ATE_Notify_6 == "true")
                        cln_LineNotificationListUSBStorage.Add(new CbtLineNotificationApi(CbtLineNotificationApi.dic_LineAuth[6]));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }

            }
            else
            {
                cln_LineNotificationListUSBStorage.Clear();
            }
        }


        #endregion



    }

    /// <Common GUI Test structure>
    /// ************************************************************************
    /// ---------------------- Common GUI Test structure ----------------------
    /// ************************************************************************
    /// </Common GUI Test structure>
    /// 


}
