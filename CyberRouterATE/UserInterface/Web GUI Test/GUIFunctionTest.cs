
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


namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        string s_ReportFileNameGUI = string.Empty;
        string s_ReportFileNameWebGUIGUI = string.Empty;
        string s_ReportPathGUI = System.Windows.Forms.Application.StartupPath + @"\report";
        string s_ReportFilePathGUI;
        string s_ReportFilePathWebGUIGUI;
        string s_ConsoleTextGUI = string.Empty;
        string s_InfoStrGUI;
        string s_BrowserDriverProcessesNameGUI = string.Empty;
        string s_FolderNameGUI = "EA8500GUI_DATE";
        string s_LogFileGUI = string.Empty;

        //bool bWebGUISingleScriptItemRunning = false;
        Thread thread_WebGUIScriptItemGUI;
        CBT_SeleniumApi.BrowserType csbt_TestBrowserGUI;
        //string sTestBrowser = string.Empty;
        List<CBT_SeleniumApi.BrowserType> csbt_TestBrowserListGUI = new List<CBT_SeleniumApi.BrowserType>();
        //List<string> sTestBrowserList = new List<string>();
        ExcelObject st_ExcelObjectGUI = new ExcelObject();

        ModelGroupStruct mgs_ModelParameterGUI = new ModelGroupStruct();
        LoginSettingGroupStruct lsgs_GatewayLoginSettingParameterGUI = new LoginSettingGroupStruct();
        LoginSetting ls_LoginSettingParametersGUI = new LoginSetting();
        TestBrowerSetting tbs_TestBrowerParametersGUI = new TestBrowerSetting();

        int i_TestRunGUI = 1;
        int i_FinalScriptIndexGUI = 0;
        int i_StepsScriptIndexGUI = 0;
        int i_duringStartStopGUI = 0;
        int i_TestScriptIndexGUI = 0;
        int i_DATA_FinalROWGUI = 16;
        int i_DATA_ROWGUI = 15;
        int i_RetryCountGUI = 1;
        int i_CommandDataIndexGUI = 0;
        string s_CurrentURLGUI = string.Empty;
        Stopwatch sw_TestTimerGUI = new Stopwatch();
        Stopwatch sw_WebGUITestTimerGUI = new Stopwatch();

        CBT_SeleniumApi cs_BrowserGUI = null;
        Comport cp_ComPortGUI = null;

        FinalScriptDataGUI[] st_ReadFinalScriptDataGUI;
        StepsScriptDataGUI[] st_ReadStepsScriptDataGUI;
        ScriptDataGUI[] st_ReadScriptDataGUI;

        //**********************************************************************************//
        //------------------ GUI Function Test Event -------------------//
        //**********************************************************************************//
        #region GUI Test Event
        private void GUITestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //TestItem = TestItemConstants.TESTITEM_THROUGHPUT;
            ToggleToolStripMenuItem(false);
            tabControl_RouterStartPage.Hide();
            webGUITestToolStripMenuItem.Checked = true;
            gUITestToolStripMenuItem.Checked = true;

            HideAllTabPage();

            this.tpGUIFunctionTest.Parent = this.tabControl_GUI;
            tabControl_GUI.Show();
            tsslMessage.Text = tabControl_GUI.TabPages[tabControl_GUI.SelectedIndex].Text + " Control Panel";

            txtGUIFunctionTestFinalScriptFilePath.Text = System.Windows.Forms.Application.StartupPath + "\\testCondition\\GUIFinal.xlsm";
            txtGUIFunctionTestStepsScriptFilePath.Text = System.Windows.Forms.Application.StartupPath + "\\testCondition\\GUISteps.xlsx";
            txtGUIFunctionTestScriptFilePath.Text = System.Windows.Forms.Application.StartupPath + "\\testCondition\\DutWebScript\\EA8500_WebGuiScript.xlsx";

            string xmlFile = System.Windows.Forms.Application.StartupPath + "\\config\\" + Constants.XML_ROUTER_GUI;
            Debug.WriteLine(File.Exists(xmlFile) ? "File exist" : "File is not existing");
        }

        private void tabControl_GUI_Selected(object sender, TabControlEventArgs e)
        {
            //ElementVisibleSetting();
            if (tabControl_GUI.SelectedIndex >= 0)
            {
                tsslMessage.Text = tabControl_GUI.TabPages[tabControl_GUI.SelectedIndex].Text + " Control Panel";
            }
        }

        private void btnGUIFunctionTestFinalScriptBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogFinal = new OpenFileDialog();

            openFileDialogFinal.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            openFileDialogFinal.Filter = "Excel file|*.xlsx|Excel file|*.xlsm";

            if (openFileDialogFinal.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialogFinal.FileName != "")
            {
                txtGUIFunctionTestFinalScriptFilePath.Clear();
                txtGUIFunctionTestFinalScriptFilePath.Text = openFileDialogFinal.FileName;
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

        private void btnGUIFunctionTestStepsScriptBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogSteps = new OpenFileDialog();

            openFileDialogSteps.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            openFileDialogSteps.Filter = "Excel file|*.xlsx";

            if (openFileDialogSteps.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialogSteps.FileName != "")
            {
                txtGUIFunctionTestStepsScriptFilePath.Clear();
                txtGUIFunctionTestStepsScriptFilePath.Text = openFileDialogSteps.FileName;
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

        private void btnGUIFunctionTestScriptBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogWebGUI = new OpenFileDialog();

            openFileDialogWebGUI.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            // Set filter for file extension and default file extension
            openFileDialogWebGUI.Filter = "Excel file|*.xlsx";

            // If the file name is not an empty string open it for opening.
            if (openFileDialogWebGUI.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialogWebGUI.FileName != "")
            {
                txtGUIFunctionTestScriptFilePath.Clear();
                txtGUIFunctionTestScriptFilePath.Text = openFileDialogWebGUI.FileName;
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

        private void btnGUIFunctionTestSave_Click(object sender, EventArgs e)
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
                File.WriteAllText(saveDebugFileDialog.FileName, txtGUIFunctionTestInformation.Text);
            }
        }

        private void btnGUIFunctionTestRun_Click(object sender, EventArgs e)
        {
            btnGUIFunctionTestRun.Enabled = false;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;


            if (bCommonFTThreadRunning == false && bCommonTestComplete == true)
            {
                SaveModelParameterGUI();
                SaveGatewayLoginParameterGUI();
                SaveTestBrowserParameterGUI();

                if (!CheckNeededParameterGUI())
                {
                    Thread.Sleep(1000);
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    btnGUIFunctionTestRun.Enabled = true;
                    return;
                }


                //-------------------------------------------------//
                //----------------- Start Test --------------------//
                //-------------------------------------------------//

                bCommonFTThreadRunning = true;
                bCommonTestComplete = false;
                ToggleFunctionTestControllerGUI(false);  /* Disable all controller */
                txtGUIFunctionTestInformation.Clear();

                //--------------- Start Test Thread ---------------//
                SetText("-------------- Start GUI ATE Test --------------", txtGUIFunctionTestInformation);
                threadCommonFT = new Thread(new ThreadStart(DoFunctionTestGUI));
                threadCommonFT.Name = "";
                threadCommonFT.Start();

                //*** Use another thread to catch the stop event of test thread ***//
                threadCommonFTstopEvent = new Thread(new ThreadStart(threadFunctionTestCatchStopEventGUI));
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
            btnGUIFunctionTestRun.Enabled = true;  //也讓Stop可以點選

        }


        #endregion

        //**********************************************************************************//
        //------------------ GUI Function Test Module ------------------//
        //**********************************************************************************//
        #region GUI Test Module


        private bool CheckNeededParameterGUI()
        {
            if (masktxtGUIFunctionTestGatewayIP.Text == "" || masktxtGUIFunctionTestGatewayIP.Text == null)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Set IP!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (nudGUIFunctionTestDefaultSettingsHTTPport.Text == "" || nudGUIFunctionTestDefaultSettingsHTTPport.Text == null)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Set HTTP port!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (txtGUIFunctionTestFinalScriptFilePath.Text == "" || txtGUIFunctionTestFinalScriptFilePath.Text == null)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Select the Final Script!!", "Warning", MessageBoxButtons.OK);
                return false;
            }
            else if (!System.IO.File.Exists(txtGUIFunctionTestFinalScriptFilePath.Text))
            {
                MessageBox.Show(new Form { TopMost = true }, "Final Script file does not exist!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (txtGUIFunctionTestStepsScriptFilePath.Text == "" || txtGUIFunctionTestStepsScriptFilePath.Text == null)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Select the Steps Script!!", "Warning", MessageBoxButtons.OK);
                return false;
            }
            else if (!System.IO.File.Exists(txtGUIFunctionTestStepsScriptFilePath.Text))
            {
                MessageBox.Show(new Form { TopMost = true }, "Steps Script file does not exist!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (txtGUIFunctionTestScriptFilePath.Text == "" || txtGUIFunctionTestScriptFilePath.Text == null)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Select the Test Script!!", "Warning", MessageBoxButtons.OK);
                return false;
            }
            else if (!System.IO.File.Exists(txtGUIFunctionTestScriptFilePath.Text))
            {
                MessageBox.Show(new Form { TopMost = true }, "WebGUI Script file does not exist!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (!CheckTestBrowserGUI())
            {
                MessageBox.Show(new Form { TopMost = true }, "Please Select the Test Browser!!", "Warning", MessageBoxButtons.OK);
                return false;
            }

            if (ckboxGUIFunctionTestConsoleLog.Checked == true)
            {
                cp_ComPortGUI = new Comport();

                if (cp_ComPortGUI.isOpen() == false)
                {
                    MessageBox.Show(new Form { TopMost = true }, "COM Port is not Open!!", "Warning", MessageBoxButtons.OK);
                    return false;
                }
            }
            return true;
        }

        private void threadFunctionTestCatchStopEventGUI()
        {
            while (bCommonFTThreadRunning == true)
            {
                //Thread.Sleep(500);
            }

            TestThreadFinishedGUI();
            threadCommonFTstopEvent.Abort();
        }

        private void TestThreadFinishedGUI()
        {
            this.Invoke(new showCommonGUIDelegate(ToggleFunctionTestGUI));
            Invoke(new SetTextCallBack(SetText), new object[] { "--------------- EA8500 GUI ATE Test End--------------", txtGUIFunctionTestInformation });
        }

        private void DoFunctionTestGUI()
        {
            WriteDebugMsg = new System.IO.StreamWriter(sDebugFilePath);
            WriteConsoleLog = new System.IO.StreamWriter(sConsoleLogFilePath);
            int iTestRun = Convert.ToInt32(tbs_TestBrowerParametersGUI.TestRun);
            int iTestTimeOut; // = Convert.ToInt32(tbsTestBrowerParameters.TestTimeOut) * 1000;

            s_InfoStrGUI = string.Format("Loading Final Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
            loadFinalScriptGUI();
            s_InfoStrGUI = string.Format("Loading Steps Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
            loadStepsScriptGUI();
            s_InfoStrGUI = string.Format("Loading WebGUI Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
            loadWebGUIScriptGUI();
            for (i_TestRunGUI = 1; i_TestRunGUI <= iTestRun; i_TestRunGUI++)
            {
                foreach (CBT_SeleniumApi.BrowserType csbtBrowserStr in csbt_TestBrowserListGUI)
                {
                    i_TestScriptIndexGUI = 0;
                    i_DATA_FinalROWGUI = 16;
                    i_DATA_ROWGUI = 15;
                    s_CurrentURLGUI = string.Empty;

                    csbt_TestBrowserGUI = csbtBrowserStr;
                    s_FolderNameGUI = "EA8500GUI_DATE";
                    CreateNewExcelFileGUI();
                    CreateNewWebGUIExcelFileGUI();

                    s_InfoStrGUI = string.Format("<< Test by Bowser: {0}   Test Run:{1} >>\n", csbtBrowserStr, i_TestRunGUI);
                    Invoke(new SetTextCallBack(SetTextC), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });

                    if (ckboxGUIFunctionTestConsoleLog.Checked == true)
                    {
                        s_InfoStrGUI = string.Format("------------------------------------- << Test by Bowser: {0}   Test Run:{1} >> -------------------------------------", csbtBrowserStr, i_TestRunGUI);
                        WriteLogTitleGUI(WriteConsoleLog, s_InfoStrGUI, "=");
                    }
                    for (i_FinalScriptIndexGUI = 0; i_FinalScriptIndexGUI < st_ReadFinalScriptDataGUI.Length; i_FinalScriptIndexGUI++)
                    {
                        //tempLogColumn = 8;
                        i_RetryCountGUI = 1;

                        s_LogFileGUI = s_ReportPathGUI + @"\" + s_FolderNameGUI + @"\" + "ConsoleLog_Index" + st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestItem + "_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".txt";

                        s_InfoStrGUI = string.Format("========Index {0} Start=========\n", st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestItem);
                        Invoke(new SetTextCallBack(SetTextC), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
                        s_InfoStrGUI = string.Format("Function Name: {0}, Name: {1}", st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].FunctionName, st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].Name);
                        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
                        s_InfoStrGUI = string.Format("Sub-Index: Start: {0}, Stop: {1}", st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].StartIndex, st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].StopIndex);
                        Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });

                        for (i_duringStartStopGUI = Convert.ToInt32(st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].StartIndex); i_duringStartStopGUI <= Convert.ToInt32(st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].StopIndex); i_duringStartStopGUI++)
                        {
                            for (i_StepsScriptIndexGUI = 0; i_StepsScriptIndexGUI < st_ReadStepsScriptDataGUI.Length; i_StepsScriptIndexGUI++)
                            {
                                if (Convert.ToInt32(st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Index) == i_duringStartStopGUI)
                                {
                                    s_InfoStrGUI = string.Format("\r\n=== Run Step-Index {0} Start===", st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Index);
                                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });

                                    //----------------------------------------------------------------//
                                    //----------------------- Main Test Thread -----------------------//
                                    //----------------------------------------------------------------//
                                    sw_TestTimerGUI.Reset();
                                    sw_TestTimerGUI.Start();
                                    bCommonSingleScriptItemRunning = true;
                                    threadCommonFTscriptItem = new Thread(new ThreadStart(FunctionTestMainFunctionGUI));
                                    //threadCommonFTscriptItem.Name = "";
                                    threadCommonFTscriptItem.Start();
                                    int timeoutNumber = 0;
                                    if (st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].FunctionName == "WebGUI")
                                    {
                                        iTestTimeOut = 20 * 60 * 1000;
                                    }
                                    else if (Int32.TryParse(st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].TestTimeOut, out timeoutNumber) == true)
                                    {
                                        iTestTimeOut = Convert.ToInt32(st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].TestTimeOut) * 1000;
                                    }
                                    else
                                    {
                                        iTestTimeOut = Convert.ToInt32(nudGUIFunctionTestDefaultTimeout.Value) * 1000;
                                    }
                                    while (sw_TestTimerGUI.ElapsedMilliseconds <= iTestTimeOut && bCommonSingleScriptItemRunning == true)
                                    {
                                        //Thread.Sleep(500);
                                    }
                                    sw_TestTimerGUI.Stop();
                                    string ElapsedTestTime = sw_TestTimerGUI.ElapsedMilliseconds.ToString();


                                    //----------------------------------------------//
                                    //------------------ Time Out ------------------//
                                    //----------------------------------------------//
                                    #region Time Out

                                    if (sw_TestTimerGUI.ElapsedMilliseconds > iTestTimeOut)
                                    {
                                        try
                                        {
                                            thread_WebGUIScriptItemGUI.Abort();
                                        }
                                        catch { }

                                        threadCommonFTscriptItem.Abort();
                                        st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestResult = "FAIL";


                                    }
                                    #endregion
                                }
                            }
                        }
                        if (st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestResult == "")
                        {
                            st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestResult = "PASS";
                        }
                        WriteTestReportGUI(true, st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI]);
                        File.AppendAllText(s_LogFileGUI, txtGUIFunctionTestInformation.Text);
                        try
                        {
                            cs_BrowserGUI.Close_WebDriver();
                        }
                        catch { }
                    }
                    TestFinishedWebGUIActionGUI();
                    TestFinishedActionGUI();
                }
            }

            bCommonFTThreadRunning = false;
        }

        private void WebGUIMainFunctionGUI()
        {
            Thread.Sleep(2000);
            CBT_SeleniumApi.GuiScriptParameter ScriptPara = new CBT_SeleniumApi.GuiScriptParameter();
            string sExceptionInfo = string.Empty;
            string sLoginURL = string.Format("http://{0}:{1}", ls_LoginSettingParametersGUI.GatewayIP, ls_LoginSettingParametersGUI.HTTP_Port);
            //string sCurrentURL = string.Empty;
            //int DATA_ROW = 15;
            bool bTestResult = true;
            int j = i_TestScriptIndexGUI;

            ScriptPara.Procedure = st_ReadScriptDataGUI[j].Procedure;
            ScriptPara.Index = st_ReadScriptDataGUI[j].TestIndex;
            ScriptPara.Action = st_ReadScriptDataGUI[j].Action;
            ScriptPara.ActionName = st_ReadScriptDataGUI[j].ActionName;
            ScriptPara.ElementType = st_ReadScriptDataGUI[j].ElementType;
            ScriptPara.ElementXpath = st_ReadScriptDataGUI[j].ElementXpath;
            ScriptPara.ElementXpath = ScriptPara.ElementXpath.Replace('\"', '\'');
            ScriptPara.RadioBtnExpectedValueXpath = st_ReadScriptDataGUI[j].RadioButtonExpectedValueXpath.Replace('\"', '\'');
            ScriptPara.WriteValue = st_ReadScriptDataGUI[j].WriteExpectedValue;
            ScriptPara.ExpectedValue = st_ReadScriptDataGUI[j].WriteExpectedValue;
            ScriptPara.URL = sLoginURL + st_ReadScriptDataGUI[j].WriteExpectedValue;
            ScriptPara.TestTimeOut = st_ReadScriptDataGUI[j].TestTimeOut;
            ScriptPara.Note = string.Empty;

            //---------------------------------------//
            //-------------- Go To URL --------------//
            //---------------------------------------//
            if (j == 0)
            {
                s_CurrentURLGUI = sLoginURL;
                cs_BrowserGUI.GoToURL(sLoginURL);
                Thread.Sleep(1000);
            }
            if (ScriptPara.Action.CompareTo("Goto") == 0 && s_CurrentURLGUI.CompareTo(ScriptPara.URL) != 0)
            {
                #region Go To URL
                s_CurrentURLGUI = ScriptPara.URL;
                Thread.Sleep(1000);
                cs_BrowserGUI.GoToURL(s_CurrentURLGUI);
                #endregion
            }

            //---------------------------------------//
            //-------------- Set Value --------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("Set") == 0)
            {
                #region Set Value
                s_InfoStrGUI = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Set Value", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });

                sw_TestTimerGUI.Stop();
                sw_WebGUITestTimerGUI.Stop();
                bool checkXPathResult = false;
                Stopwatch waitingXpath = new Stopwatch();
                waitingXpath.Start();
                do
                {
                    checkXPathResult = cs_BrowserGUI.CheckXPathDisplayed(ScriptPara.ElementXpath);
                } while (waitingXpath.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);
                waitingXpath.Reset();
                sw_TestTimerGUI.Start();
                sw_WebGUITestTimerGUI.Start();
                if (checkXPathResult == true)
                {
                    Thread.Sleep(2000);

                    try
                    {
                        if (ScriptPara.WriteValue != "" && st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].WriteValue[i_CommandDataIndexGUI] != "" && st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].WriteValue[i_CommandDataIndexGUI] != null)
                        {
                            ScriptPara.WriteValue = st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].WriteValue[i_CommandDataIndexGUI];
                            i_CommandDataIndexGUI++;
                        }
                    }
                    catch { }
                    ScriptPara.Note = string.Empty;
                    try
                    {
                        cs_BrowserGUI.SetWebElementValue(ref ScriptPara);
                    }
                    catch
                    {
                        ExceptionActionGUI(s_CurrentURLGUI, ref ScriptPara);
                        ScriptPara.Note = "Set Value Error:\n" + ScriptPara.Note;
                        SwitchToRGconsoleGUI();
                        WriteconsoleLogGUI();
                        //return false;
                    }
                }
                else if (checkXPathResult == false)
                {
                    s_InfoStrGUI = string.Format("...Couldn't find the element!");
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
                    ExceptionActionGUI(s_CurrentURLGUI, ref ScriptPara);
                    ScriptPara.Note = "Set Value Error:\n" + ScriptPara.Note;
                    SwitchToRGconsoleGUI();
                    WriteconsoleLogGUI();
                }
                
                #endregion
            }

            //---------------------------------------//
            //-------------- Get Value --------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("Get") == 0)
            {
                #region Get Value
                s_InfoStrGUI = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Get Value", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });

                ScriptPara.Note = string.Empty;
                ScriptPara.GetValue = string.Empty;

                if (ScriptPara.ElementType.CompareTo("RADIO_BUTTON") == 0)
                {
                    ScriptPara.ElementXpath = st_ReadScriptDataGUI[j].RadioButtonExpectedValueXpath.Replace('\"', '\'');
                }

                try
                {
                    bTestResult = cs_BrowserGUI.GetWebElementValue(ref ScriptPara); // Set Value
                }
                catch
                {
                    ExceptionActionGUI(s_CurrentURLGUI, ref ScriptPara);
                    ScriptPara.Note = "Execute SubmitButton Error:\n" + ScriptPara.Note;
                    //return false;
                }


                //---------- Write Test Report ----------//
                //WriteTestReportGUI(j, TestResult, ref DATA_ROW, ScriptPara.Note);
                #endregion
            }

            //---------------------------------------//
            //--------------- ReLogin ---------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("ReLogin") == 0)
            {
                #region ReLogin
                Invoke(new SetTextCallBack(SetText), new object[] { "Log in again...", txtGUIFunctionTestInformation });
                Re_LoginGUI();
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
                Invoke(new SetTextCallBack(SetText), new object[] { "Login...", txtGUIFunctionTestInformation });
                while (true)
                {
                    if (pingTimer.ElapsedMilliseconds > (Convert.ToInt32(120000)))
                    {
                        break;
                    }

                    if (PingClient(ls_LoginSettingParametersGUI.GatewayIP, 1000))
                    {
                        Invoke(new SetTextCallBack(SetText), new object[] { "\tPing Successfully!!", txtGUIFunctionTestInformation });
                        break;
                    }

                    Thread.Sleep(1000);
                    string Info = String.Format(".");
                    Invoke(new SetTextCallBack(SetText), new object[] { Info, txtGUIFunctionTestInformation });
                }
                Thread.Sleep(3000);
                String[] splitLoginInfo = ScriptPara.WriteValue.Split('/');
                cs_BrowserGUI.loginAlertMessage(ScriptPara.URL, splitLoginInfo[0], splitLoginInfo[1]);
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
                s_InfoStrGUI = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:File Upload", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });

                Thread.Sleep(3000);
                ScriptPara.Note = string.Empty;
                try
                {
                    cs_BrowserGUI.fileUploadE8350(ref ScriptPara);
                }
                catch
                {
                    ExceptionActionGUI(s_CurrentURLGUI, ref ScriptPara);
                    ScriptPara.Note = "Set Value Error:\n" + ScriptPara.Note;
                    SwitchToRGconsoleGUI();
                    WriteconsoleLogGUI();
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
                    s_InfoStrGUI = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Waiting for {2} seconds", ScriptPara.Index, ScriptPara.ActionName, ScriptPara.TestTimeOut);
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
                    sw_TestTimerGUI.Stop();
                    Thread.Sleep(Convert.ToInt16(ScriptPara.TestTimeOut) * 1000);
                    sw_TestTimerGUI.Start();
                    SwitchToRGconsoleGUI();
                    WriteconsoleLogGUI();
                }
                else
                {
                    s_InfoStrGUI = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Waiting for the XPath, it won't be more than {2} seconds", ScriptPara.Index, ScriptPara.ActionName, ScriptPara.TestTimeOut);
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
                    sw_TestTimerGUI.Stop();
                    bool checkXPathResult = false;
                    Stopwatch waitingXpath = new Stopwatch();
                    waitingXpath.Start();
                    do
                    {
                        checkXPathResult = cs_BrowserGUI.CheckXPathDisplayed(ScriptPara.ElementXpath);
                    } while (waitingXpath.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);
                    waitingXpath.Reset();
                    sw_TestTimerGUI.Start();
                    SwitchToRGconsoleGUI();
                    WriteconsoleLogGUI();
                }
                #endregion
            }

            //---------------------------------------//
            //--------------- Poke ---------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("Poke") == 0)
            {
                #region Set Value
                s_InfoStrGUI = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Poke", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });

                sw_TestTimerGUI.Stop();
                sw_WebGUITestTimerGUI.Stop();
                bool checkXPathResult = false;
                Stopwatch waitingXpath = new Stopwatch();
                waitingXpath.Start();
                do
                {
                    checkXPathResult = cs_BrowserGUI.CheckXPathDisplayed(ScriptPara.ElementXpath);
                } while (waitingXpath.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);
                waitingXpath.Reset();
                sw_TestTimerGUI.Start();
                sw_WebGUITestTimerGUI.Start();
                if (checkXPathResult == true)
                {
                    Thread.Sleep(2000);

                    try
                    {       //So that the StepsScript could set parameters such as "test1, test2...." 
                        if (ScriptPara.WriteValue != "" && st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].WriteValue[i_CommandDataIndexGUI] != "" && st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].WriteValue[i_CommandDataIndexGUI] != null)
                        {
                            ScriptPara.WriteValue = st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].WriteValue[i_CommandDataIndexGUI];
                            i_CommandDataIndexGUI++;
                        }
                    }
                    catch { }
                    ScriptPara.Note = string.Empty;
                    try
                    {
                        cs_BrowserGUI.SetWebElementValue(ref ScriptPara);
                    }
                    catch
                    {
                        ScriptPara.Note = "Didn't poke";
                    }
                }
                else if (checkXPathResult == false)
                {
                    s_InfoStrGUI = string.Format("Didn't poke");
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
                    ScriptPara.Note = "Didn't poke";
                }

                #endregion
            }

            //---------------------------------------//
            //--------------- HoldMenu --------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("HoldMenu") == 0)
            {
                #region HoldMenu
                s_InfoStrGUI = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Hold Menu", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });

                Thread.Sleep(3000);
                ScriptPara.Note = string.Empty;
                try
                {
                    cs_BrowserGUI.holdMenu(ref ScriptPara);
                }
                catch
                {
                    ExceptionActionGUI(s_CurrentURLGUI, ref ScriptPara);
                    ScriptPara.Note = "Set Value Error:\n" + ScriptPara.Note;
                    SwitchToRGconsoleGUI();
                    WriteconsoleLogGUI();
                }
                #endregion
            }

            //---------------------------------------//
            //--------------- CloseDriver ---------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("CloseDriver") == 0)
            {
                #region CloseDriver
                s_InfoStrGUI = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Close Driver", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
                cs_BrowserGUI.Close_WebDriver();
                #endregion
            }

            //---------------------------------------//
            //--------------- OpenDriver ---------------//
            //---------------------------------------//
            else if (ScriptPara.Action.CompareTo("OpenDriver") == 0)
            {
                #region OpenDriver
                s_InfoStrGUI = string.Format("**      [WebGUI Index]: {0}    [Action Name]:{1}    [Action]:Open Driver", ScriptPara.Index, ScriptPara.ActionName);
                Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
                TestInitialGUI();
                #endregion
            }


            WriteconsoleLogGUI();



            string Key1 = "Login-";
            string Key2 = "ReLogin";


            if (ScriptPara.ActionName.ToLower().IndexOf(Key1.ToLower()) < 0 && ScriptPara.ActionName.ToLower().IndexOf(Key2.ToLower()) < 0)
            {
                //---------- Write Test Report ----------//
                //WriteTestReportGUI(j, TestResult, ref DATA_ROW, ScriptPara.Note);
                WriteWebGUITestReportGUI(bTestResult, ScriptPara);
            }

            if (ScriptPara.Note != string.Empty)
            {
                st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestResult = "FAIL";
                st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].Comment = "FAIL in: \n" + st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Name;
                i_duringStartStopGUI = Convert.ToInt32(st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].StopIndex) + 1;
            }
            bWebGUISingleScriptItemRunning = false;
        }
        
        //private static extern bool SetForegroundWindow(IntPtr hWnd);
        private void FunctionTestMainFunctionGUI()
        {
            switch (st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].FunctionName)
            {
                case "WebGUI":
                    //----------------------------------------------------------------//
                    //---------------------------- WebGUI ----------------------------//
                    //----------------------------------------------------------------//
                    #region WebGUI
                    s_InfoStrGUI = string.Format("** [Step Index]: {0}   [Function Name]:WebGUI    [Name]:{1} ", st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Index, st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Name);
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
                    i_CommandDataIndexGUI = 0;
                    for (i_TestScriptIndexGUI = 0; i_TestScriptIndexGUI < st_ReadScriptDataGUI.Length; i_TestScriptIndexGUI++)
                    {
                        if (st_ReadScriptDataGUI[i_TestScriptIndexGUI].Procedure == st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Action)
                        {
                            if (st_ReadScriptDataGUI[i_TestScriptIndexGUI].TestIndex.CompareTo("") != 0 && st_ReadScriptDataGUI[i_TestScriptIndexGUI].Action.CompareTo("") != 0)
                            {
                                bWebGUISingleScriptItemRunning = true;
                                int iWebGUITestTimeOut = 120 * 1000;
                                if (i_TestScriptIndexGUI == 0)
                                {
                                    TestInitialGUI();
                                }
                                if (bWebGUISingleScriptItemRunning == true)
                                {
                                    //Process[] mainForm = System.Diagnostics.Process.GetProcessesByName("CyberRouterATE");
                                    //SetForegroundWindow(mainForm[0].Handle);
                                    sw_WebGUITestTimerGUI.Reset();
                                    sw_WebGUITestTimerGUI.Start();
                                    thread_WebGUIScriptItemGUI = new Thread(new ThreadStart(WebGUIMainFunctionGUI));
                                    thread_WebGUIScriptItemGUI.Start();
                                    int iTimeOutNull = 0;
                                    if (Int32.TryParse(st_ReadScriptDataGUI[i_TestScriptIndexGUI].TestTimeOut, out iTimeOutNull) == true)
                                    {
                                        iWebGUITestTimeOut = Convert.ToInt32(st_ReadScriptDataGUI[i_TestScriptIndexGUI].TestTimeOut) * 1000;
                                    }
                                    while (sw_WebGUITestTimerGUI.ElapsedMilliseconds <= iWebGUITestTimeOut && bWebGUISingleScriptItemRunning == true)
                                    {
                                        //Thread.Sleep(500);
                                    }
                                    sw_WebGUITestTimerGUI.Stop();


                                    //----------------------------------------------//
                                    //--------------- WebGUI Time Out --------------//
                                    //----------------------------------------------//
                                    #region Time Out
                                    if (sw_WebGUITestTimerGUI.ElapsedMilliseconds > iWebGUITestTimeOut)
                                    {
                                        thread_WebGUIScriptItemGUI.Abort();
                                        bWebGUISingleScriptItemRunning = false;

                                        string TestStatus = string.Empty;
                                        if (!CheckExtraTestStatusGUI(ref TestStatus))
                                        {
                                            TestStatus = "Time Out";
                                        }
                                        else
                                        {
                                            TestStatus = string.Format("Time Out. " + TestStatus);
                                        }

                                        WriteWebGUITestReportGUI(i_TestScriptIndexGUI, false, TestStatus);

                                        if (StopWhenTestErrorGUI())
                                        {
                                            TestFinishedWebGUIActionGUI();
                                            //TestFinishedActionGUI();
                                            Invoke(new SetTextCallBack(SetText), new object[] { "Test Abort.", txtGUIFunctionTestInformation });
                                            bCommonFTThreadRunning = false;
                                            return;
                                        }
                                        else
                                        {
                                            if (i_RetryCountGUI >= 1)
                                            {
                                                //------ Stop Item Test and Close Web Driver------//
                                                cs_BrowserGUI.Close_WebDriver();
                                                Thread.Sleep(5000);

                                                //-------- Initial and Open New Web Driver--------//
                                                //TestInitialGUI();
                                                string Info = string.Format("{0}!! Login again...", TestStatus);
                                                Invoke(new SetTextCallBack(SetText), new object[] { Info, txtGUIFunctionTestInformation });
                                                i_duringStartStopGUI = Convert.ToInt32(st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].StartIndex);
                                                i_StepsScriptIndexGUI = -1;
                                                bCommonSingleScriptItemRunning = false;
                                                i_RetryCountGUI--;
                                                return;
                                                //Re_LoginGUI();
                                            }
                                            else
                                            {
                                                st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestResult = "ERROR";
                                                st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].Comment = "ERROR in: \n" + st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Name;
                                                i_duringStartStopGUI = Convert.ToInt32(st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].StopIndex) + 1;

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
                                //WebGUIMainFunctionGUI();
                            }
                        }

                        if (st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestResult == "FAIL" || st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestResult == "ERROR")
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
                    s_InfoStrGUI = string.Format("** [Step Index]: {0}   [Function Name]:CMD    [Name]:{1} ", st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Index, st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Name);
                    Invoke(new SetTextCallBack(SetText), new object[] { s_InfoStrGUI, txtGUIFunctionTestInformation });
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
                    while (st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Command[iCommandDataIndex] != "" && st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Command[iCommandDataIndex] != null)
                    {
                        process.StandardInput.WriteLine(@st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Command[iCommandDataIndex]);
                        iCommandDataIndex++;
                    }
                    process.StandardInput.WriteLine("exit");
                    // Read the standard output of the process.
                    string output;
                    //string tempoutput = "ExpectedValue::";
                    while ((output = process.StandardOutput.ReadLine()) != null || (output = process.StandardError.ReadLine()) != null)
                    {
                        //tempoutput += output + "\n";
                        st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].GetValue += output + "\n";  //需改成 ].GetValue 
                        Invoke(new SetTextCallBack(SetText), new object[] { output, txtGUIFunctionTestInformation });
                    }
                    //st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, tempLogColumn] = tempoutput;
                    //tempLogColumn++;
                    //需Uncomment
                    if (st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].GetValue.Contains(st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].ExpectedValue) == false)
                    {
                        st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestResult = "FAIL";
                        st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].Comment = "FAIL in: \n" + st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].Name;
                        i_duringStartStopGUI = Convert.ToInt32(st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].StopIndex) + 1;
                    }
                    process.WaitForExit();
                    process.Close();
                    break;
                    #endregion

                default:
                    st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].TestResult = "ERROR";
                    st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].Comment = "ERROR: \nThere is no function named \"" + st_ReadStepsScriptDataGUI[i_StepsScriptIndexGUI].FunctionName + "\" in the program";
                    i_duringStartStopGUI = Convert.ToInt32(st_ReadFinalScriptDataGUI[i_FinalScriptIndexGUI].StopIndex) + 1;
                    break;
            }

            bCommonSingleScriptItemRunning = false;
            Thread.Sleep(2000);


        }

        private void ToggleFunctionTestControllerGUI(bool Toggle)
        {
            btnGUIFunctionTestRun.Text = Toggle ? "Run" : "Stop";

            //----- Function Test -----//
            gbox_GUIModel.Enabled = Toggle;
            gbox_GUIGetewaySetting.Enabled = Toggle;
            gbox_GUIFunctionTestScript.Enabled = Toggle;
            gbox_GUIBrowser.Enabled = Toggle;


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
                btnGUIFunctionTestRun.Enabled = false;
                WriteconsoleLogGUI();
                WriteConsoleLog.Close();
                WriteDebugMsg.Close();
                try
                {
                    thread_WebGUIScriptItemGUI.Abort();
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
                    SetExcelCellsHighGUI();
                    SetWebGUIExcelCellsHighGUI();
                    TestFinishedActionGUI();

                }
                catch { }

                SaveEndTime_and_IEversion_to_Excel(cs_BrowserGUI);
                Save_and_Close_Excel_File();
                File.SetAttributes(txtGUIFunctionTestFinalScriptFilePath.Text, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
                File.SetAttributes(txtGUIFunctionTestStepsScriptFilePath.Text, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
                File.SetAttributes(txtGUIFunctionTestScriptFilePath.Text, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀

                try
                {
                    cs_BrowserGUI.Close_WebDriver();
                }
                catch { }
                //Close_WebDriver_and_Browser();

                //MessageBox.Show(this, "Test complete !!!", "Information", MessageBoxButtons.OK);
                MessageBox.Show(new Form { TopMost = true }, "Test complete !!!", "ATE Information", MessageBoxButtons.OK); // 將 MessageBox於桌面置頂，否則會被 Browser蓋住
                btnGUIFunctionTestRun.Enabled = true;
                bCommonTestComplete = true;
            }
        }

        private void WriteconsoleLogGUI()
        {
            if (ckboxGUIFunctionTestConsoleLog.Checked == true)
            {
                s_ConsoleTextGUI = string.Empty;

                for (int i = 0; i < 2; i++)
                {
                    s_ConsoleTextGUI += ReadConsoleLogByBytesGUI(cp_ComPortGUI);
                    Thread.Sleep(500);
                }

                s_ConsoleTextGUI += "\r\n\r\n\r\n****************************************************************************************************************\r\n\r\n\r\n";
                WriteConsoleLog.Write(s_ConsoleTextGUI);
            }
        }

        private void SetExcelCellsHighGUI()
        {
            string CELL1;
            string CELL2;

            //excelWorkSheetCommon = excelWorkBookCommon.Sheets[2];
            //CELL1 = "B8";
            //CELL2 = "B25";
            //excelRangeCommon = excelWorkSheetCommon.Range[CELL1, CELL2];
            //excelRangeCommon.RowHeight = 15;

            st_ExcelObjectGUI.excelWorkSheet = st_ExcelObjectGUI.excelWorkBook.Sheets[1];
            CELL1 = "B16";
            CELL2 = "B50000";
            st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Range[CELL1, CELL2];
            st_ExcelObjectGUI.excelRange.RowHeight = 15;

        }

        private void SetWebGUIExcelCellsHighGUI()
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

        private void ToggleFunctionTestGUI()
        {
            ToggleFunctionTestControllerGUI(true);
            Debug.WriteLine("Toggle");
        }

        private void SaveModelParameterGUI()
        {
            mgs_ModelParameterGUI.ModelName = txtGUIFunctionTestModelName.Text;
            mgs_ModelParameterGUI.SN = txtGUIFunctionTestModelSerialNumber.Text;
            mgs_ModelParameterGUI.SWver = txtGUIFunctionTestModelSwVersion.Text;
            mgs_ModelParameterGUI.HWver = txtGUIFunctionTestModelHwVersion.Text;
        }

        private void SaveGatewayLoginParameterGUI()
        {
            //lsgsGatewayLoginSettingParameter.GatewayIP = masktxtGUIFunctionTestGatewayIP.Text;

            ls_LoginSettingParametersGUI.GatewayIP = masktxtGUIFunctionTestGatewayIP.Text;
            ls_LoginSettingParametersGUI.HTTP_Port = nudGUIFunctionTestDefaultSettingsHTTPport.Value.ToString();
            //lsLoginSettingParameters.RebootWaitTime = Convert.ToInt32(nudGUIFunctionTestRebootWaitTime.Value.ToString());

            if (ckboxGUIFunctionTestConsoleLog.Checked)
                ls_LoginSettingParametersGUI.ConsoleLog = true;
            else
                ls_LoginSettingParametersGUI.ConsoleLog = false;
        }

        private void SaveTestBrowserParameterGUI()
        {
            tbs_TestBrowerParametersGUI.TestRun = nudGUIFunctionTestTestRun.Value.ToString();
            if (ckboxGUIFunctionTestStopWhenTestError.Checked)
                tbs_TestBrowerParametersGUI.StopWhenTestError = true;

        }

        private bool CheckTestBrowserGUI()
        {
            csbt_TestBrowserListGUI.Clear();

            if (ckbox_GUIFunctionTest_TestBrowser_Chrome.Checked == false &&
                ckbox_GUIFunctionTest_TestBrowser_FireFox.Checked == false &&
                ckbox_GUIFunctionTest_TestBrowser_IE.Checked == false)
            {
                return false;
            }

            if (ckbox_GUIFunctionTest_TestBrowser_Chrome.Checked == true)
                csbt_TestBrowserListGUI.Add(CBT_SeleniumApi.BrowserType.Chrome);
            if (ckbox_GUIFunctionTest_TestBrowser_FireFox.Checked == true)
                csbt_TestBrowserListGUI.Add(CBT_SeleniumApi.BrowserType.FireFox);
            if (ckbox_GUIFunctionTest_TestBrowser_IE.Checked == true)
                csbt_TestBrowserListGUI.Add(CBT_SeleniumApi.BrowserType.IE);

            return true;
        }

        private void CreateNewExcelFileGUI()
        {
            s_ReportFileNameGUI = "MODEL_FinalReport_SN_SW_HW_DATE.xlsx";
            s_ReportFileNameGUI = getReportFileNameGUI(s_ReportFileNameGUI);

            string sSubPath;
            bool bIsExists;

            s_FolderNameGUI = s_FolderNameGUI.Replace("DATE", DateTime.Now.ToString("yyyy-MMdd-HHmmss"));
            sSubPath = System.Windows.Forms.Application.StartupPath + "\\report\\" + s_FolderNameGUI;
            bIsExists = System.IO.Directory.Exists(sSubPath);

            if (!bIsExists)
                System.IO.Directory.CreateDirectory(sSubPath);

            s_ReportFilePathGUI = s_ReportPathGUI + @"\" + s_FolderNameGUI + @"\" + s_ReportFileNameGUI;
            initialExcel_CommonCaseGUI(s_ReportFilePathGUI);
        }

        private void CreateNewWebGUIExcelFileGUI()
        {
            s_ReportFileNameWebGUIGUI = "MODEL_WebGUIReport_SN_SW_HW_DATE_BROWSER.xlsx";
            s_ReportFileNameWebGUIGUI = getReportFileNameGUI(s_ReportFileNameWebGUIGUI);

            s_ReportFilePathWebGUIGUI = s_ReportPathGUI + @"\" + s_FolderNameGUI + @"\" + s_ReportFileNameWebGUIGUI;
            initialWebGUIExcel_CommonCaseGUI(s_ReportFilePathWebGUIGUI);
        }

        public string getReportFileNameGUI(string reportTargetfileName)
        {
            reportTargetfileName = reportTargetfileName.Replace("MODEL", (mgs_ModelParameterGUI.ModelName == "") ? "GUI" : txtGUIFunctionTestModelName.Text);
            reportTargetfileName = reportTargetfileName.Replace("SN", mgs_ModelParameterGUI.SN);
            reportTargetfileName = reportTargetfileName.Replace("SW", mgs_ModelParameterGUI.SWver);
            reportTargetfileName = reportTargetfileName.Replace("HW", mgs_ModelParameterGUI.HWver);
            reportTargetfileName = reportTargetfileName.Replace("DATE", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            reportTargetfileName = reportTargetfileName.Replace("BROWSER", Convert.ToString(csbt_TestBrowserGUI));
            return reportTargetfileName;
        }

        private void initialExcel_CommonCaseGUI(string filePath)
        {
            st_ExcelObjectGUI.excelApp = new Excel.Application();

            /*** Set Excel visible ***/
            st_ExcelObjectGUI.excelApp.Visible = true;

            /*** Do not show alert ***/
            st_ExcelObjectGUI.excelApp.DisplayAlerts = false;

            st_ExcelObjectGUI.excelApp.UserControl = true;
            st_ExcelObjectGUI.excelApp.Interactive = false;

            //Set font and font size attributes
            st_ExcelObjectGUI.excelApp.StandardFont = "Times New Roman";
            st_ExcelObjectGUI.excelApp.StandardFontSize = 11;

            /*** This method is used to open an Excel workbook by passing the file path as a parameter to this method. ***/
            st_ExcelObjectGUI.excelWorkBook = st_ExcelObjectGUI.excelApp.Workbooks.Add(misValue);

            st_ExcelObjectGUI.excelWorkSheet = st_ExcelObjectGUI.excelWorkBook.Sheets[1];
            st_ExcelObjectGUI.excelWorkSheet.Name = "Test Results";
            createExcelTitleGUI();

            saveExcelGUI(filePath);
        }

        private void initialWebGUIExcel_CommonCaseGUI(string filePath)
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
            createWebGUIExcelTitleGUI();

            saveWebGUIExcelGUI(filePath);
        }

        private void createExcelTitleGUI()
        {
            //*** SettingExcelFont() ***//
            //--- [ColorIndex] 黑:1, 紅:3, 藍:5,淡黃:27, 淡橘:44 
            //--- [Font Style] U:底線 B:粗體 I:斜體 (可同時設定多個屬性)   ex:"UBI"、"BI"、"I"、"B"...

            //st_ExcelObjectGUI.excelApp.Cells[2, 2] = "CyberATE EA8500 GUI Test Report";
            st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Cells[2, 2];
            SettingExcelFontGUI(st_ExcelObjectGUI.excelRange, "Times New Roman", 14, "U", 1);
            st_ExcelObjectGUI.excelWorkSheet.Cells[2, 2] = "CyberATE EA8500 GUI Final Report";

            st_ExcelObjectGUI.excelApp.Cells[4, 2] = "Test";
            st_ExcelObjectGUI.excelApp.Cells[5, 2] = "Station";
            st_ExcelObjectGUI.excelApp.Cells[6, 2] = "Start Time";
            st_ExcelObjectGUI.excelApp.Cells[7, 2] = "End Time";
            st_ExcelObjectGUI.excelApp.Cells[8, 2] = "Model";
            st_ExcelObjectGUI.excelApp.Cells[9, 2] = "Serial";
            st_ExcelObjectGUI.excelApp.Cells[10, 2] = "SW Version";
            st_ExcelObjectGUI.excelApp.Cells[11, 2] = "HW Version";

            st_ExcelObjectGUI.excelApp.Cells[4, 3] = "EA8500 GUI Test";
            st_ExcelObjectGUI.excelApp.Cells[5, 3] = "CyberTAN ATE-" + txtGUIFunctionTestModelName.Text;
            st_ExcelObjectGUI.excelApp.Cells[6, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            st_ExcelObjectGUI.excelApp.Cells[8, 3] = txtGUIFunctionTestModelName.Text;
            st_ExcelObjectGUI.excelApp.Cells[9, 3] = txtGUIFunctionTestModelSerialNumber.Text;
            st_ExcelObjectGUI.excelApp.Cells[10, 3] = txtGUIFunctionTestModelSwVersion.Text;
            st_ExcelObjectGUI.excelApp.Cells[11, 3] = txtGUIFunctionTestModelHwVersion.Text;

            /* Set cells width */
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 1], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 1]].ColumnWidth = 2;
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 2], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 2]].ColumnWidth = 15; /* column B */
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 3], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 3]].ColumnWidth = 15;
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 4], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 4]].ColumnWidth = 15; /* column B */
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 5], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 5]].ColumnWidth = 15;
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 6], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 6]].ColumnWidth = 15; /* column B */
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 7], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 7]].ColumnWidth = 15;
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 8], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 8]].ColumnWidth = 15; /* column B */
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 9], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 9]].ColumnWidth = 15; /* column B */
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 10], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 10]].ColumnWidth = 15; /* column B */
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[2, 11], st_ExcelObjectGUI.excelWorkSheet.Cells[10, 11]].ColumnWidth = 15; /* column B */

            st_ExcelObjectGUI.excelApp.Cells[3, 8] = "Option";
            st_ExcelObjectGUI.excelApp.Cells[3, 9] = "Value";

            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelApp.Cells[3, 8], st_ExcelObjectGUI.excelApp.Cells[7, 9]].Borders.LineStyle = 1;

            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelApp.Cells[3, 8], st_ExcelObjectGUI.excelApp.Cells[3, 9]].Font.Underline = true;
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelApp.Cells[3, 8], st_ExcelObjectGUI.excelApp.Cells[3, 9]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelApp.Cells[3, 8], st_ExcelObjectGUI.excelApp.Cells[3, 9]].Font.FontStyle = "Bold";

            st_ExcelObjectGUI.excelWorkSheet.Cells[4, 8] = "Loop";
            st_ExcelObjectGUI.excelWorkSheet.Cells[4, 9] = i_TestRunGUI;
            st_ExcelObjectGUI.excelWorkSheet.Cells[5, 8] = "Final Script";
            st_ExcelObjectGUI.excelWorkSheet.Cells[5, 9] = txtGUIFunctionTestFinalScriptFilePath.Text;
            st_ExcelObjectGUI.excelWorkSheet.Cells[6, 8] = "Steps Script";
            st_ExcelObjectGUI.excelWorkSheet.Cells[6, 9] = txtGUIFunctionTestStepsScriptFilePath.Text;
            st_ExcelObjectGUI.excelWorkSheet.Cells[7, 8] = "WebGUI Script";
            st_ExcelObjectGUI.excelWorkSheet.Cells[7, 9] = txtGUIFunctionTestScriptFilePath.Text;

            st_ExcelObjectGUI.excelApp.Cells[15, 2] = "Item No";
            st_ExcelObjectGUI.excelApp.Cells[15, 3] = "Function Name";
            st_ExcelObjectGUI.excelApp.Cells[15, 4] = "Name";
            st_ExcelObjectGUI.excelApp.Cells[15, 5] = "Test Result";
            st_ExcelObjectGUI.excelApp.Cells[15, 6] = "Comment";
            st_ExcelObjectGUI.excelApp.Cells[15, 7] = "Log File";
            //st_ExcelObjectGUI.excelApp.Cells[15, 8] = "";

            st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelWorkSheet.Cells[14, 2], st_ExcelObjectGUI.excelWorkSheet.Cells[14, 10]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


            //st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelApp.Cells[15, 2], st_ExcelObjectGUI.excelApp.Cells[15, 10]].Font.FontStyle = "Bold";
            //st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelApp.Cells[15, 2], st_ExcelObjectGUI.excelApp.Cells[15, 10]];
            //SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "C");


            ////*** SettingExcelAlignment() ***//
            ////--- [direction] H:水平方向 V:垂直方向 (可同時設定多個屬性)   ex:"HV"、"H"、"V"
            ////--- [AlignmentType] C:置中 L:靠左 R:靠右

            ////st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelApp.Cells[6, 3], st_ExcelObjectGUI.excelApp.Cells[11, 3]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Range[st_ExcelObjectGUI.excelApp.Cells[4, 3], st_ExcelObjectGUI.excelApp.Cells[12, 3]];
            //SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "C");

            //st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Range["E4", "E11"];
            //SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "C");


            /*** Set cells width ***/
            //SetExcelCellsWidthGUI();

        }

        private void createWebGUIExcelTitleGUI()
        {
            //*** SettingExcelFont() ***//
            //--- [ColorIndex] 黑:1, 紅:3, 藍:5,淡黃:27, 淡橘:44 
            //--- [Font Style] U:底線 B:粗體 I:斜體 (可同時設定多個屬性)   ex:"UBI"、"BI"、"I"、"B"...

            //excelAppCommon.Cells[2, 2] = "CyberATE EA8500 GUI Test Report";
            excelRangeCommon = excelWorkSheetCommon.Cells[2, 2];
            SettingExcelFontGUI(excelRangeCommon, "Times New Roman", 14, "U", 1);
            excelWorkSheetCommon.Cells[2, 2] = "CyberATE EA8500 GUI WebGUI Test Report";

            excelAppCommon.Cells[4, 2] = "Test";
            excelAppCommon.Cells[5, 2] = "Station";
            excelAppCommon.Cells[6, 2] = "Start Time";
            excelAppCommon.Cells[7, 2] = "End Time";
            excelAppCommon.Cells[8, 2] = "Model";
            excelAppCommon.Cells[9, 2] = "Serial";
            excelAppCommon.Cells[10, 2] = "SW Version";
            excelAppCommon.Cells[11, 2] = "HW Version";
            //excelAppCommon.Cells[12, 2] = "Test Browser";

            excelAppCommon.Cells[4, 3] = "EA8500 GUI Test";
            excelAppCommon.Cells[5, 3] = "CyberTAN ATE-" + txtGUIFunctionTestModelName.Text;
            excelAppCommon.Cells[6, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            excelAppCommon.Cells[8, 3] = txtGUIFunctionTestModelName.Text;
            excelAppCommon.Cells[9, 3] = txtGUIFunctionTestModelSerialNumber.Text;
            excelAppCommon.Cells[10, 3] = txtGUIFunctionTestModelSwVersion.Text;
            excelAppCommon.Cells[11, 3] = txtGUIFunctionTestModelHwVersion.Text;

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
            SetExcelCellsWidthGUI();

        }

        private void SetExcelCellsWidthGUI()
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

        private void SettingExcelFontGUI(Excel.Range excelRangeCommon, string FONT_TYPE, int FONT_SIZE, string FONT_STYLE, int COLOR_INDEX)
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

        private void WriteLogTitleGUI(System.IO.StreamWriter WriteLog, string titleStr, string symbol)
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

        private void TestInitialGUI()
        {
            cs_BrowserGUI = new CBT_SeleniumApi();

            if (!cs_BrowserGUI.init(csbt_TestBrowserGUI))
            {
                SetExcelCellsHighGUI();
                SetWebGUIExcelCellsHighGUI();
                SaveEndTime_and_Version_to_ExcelGUI();
                SaveEndTime_and_Version_to_WebGUIExcelGUI();
                Save_and_Close_Excel_File();
                Save_and_Close_Excel_File_Common(st_ExcelObjectGUI);
                cs_BrowserGUI.Close_WebDriver();
                bCommonFTThreadRunning = false;
                bWebGUISingleScriptItemRunning = false;
                bCommonSingleScriptItemRunning = false;
                return;
            }
            cs_BrowserGUI.SettingTimeout(60);
            cs_BrowserGUI.WindowMaximize();
            Thread.Sleep(1000);
            cs_BrowserGUI.WindowMinimize();
            Thread.Sleep(3000);
        }

        private void SaveEndTime_and_Version_to_ExcelGUI()
        {
            try
            {
                st_ExcelObjectGUI.excelWorkSheet = st_ExcelObjectGUI.excelWorkBook.Sheets[1];

                //st_ExcelObjectGUI.excelApp.Cells[10, 4] = "Script Path";
                //st_ExcelObjectGUI.excelApp.Cells[10, 5] = txtGUIFunctionTestFinalScriptFilePath.Text;
                //st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Cells[10, 5];
                //SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "R");
                //st_ExcelObjectGUI.excelApp.Cells[11, 4] = "Browser Version";
                //st_ExcelObjectGUI.excelApp.Cells[11, 5] = cs_BrowserGUI.BrowserVersion();
                //st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Cells[11, 5];
                //SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "R");

                st_ExcelObjectGUI.excelApp.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");  // Save End Time to Excel
            }
            catch { }

        }

        private void SaveEndTime_and_Version_to_WebGUIExcelGUI()
        {
            try
            {
                excelWorkSheetCommon = excelWorkBookCommon.Sheets[1];

                excelAppCommon.Cells[10, 4] = "Script Path";
                excelAppCommon.Cells[10, 5] = txtGUIFunctionTestScriptFilePath.Text;
                excelRangeCommon = excelWorkSheetCommon.Cells[10, 5];
                SettingExcelAlignment(excelRangeCommon, "H", "R");
                excelAppCommon.Cells[11, 4] = "Browser Version";

                excelAppCommon.Cells[11, 5] = cs_BrowserGUI.BrowserVersion();

                excelRangeCommon = excelWorkSheetCommon.Cells[11, 5];
                SettingExcelAlignment(excelRangeCommon, "H", "R");



                excelAppCommon.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");  // Save End Time to Excel
            }
            catch { }

        }

        private bool CheckExtraTestStatusGUI(ref string TestStatus)
        {
            //-----------------------------------------------//
            //-------------- Check Ping Status --------------//
            //-----------------------------------------------//
            bool pingStatus = false;
            Stopwatch pingTimer = new Stopwatch();
            pingTimer.Reset();
            pingTimer.Start();
            Invoke(new SetTextCallBack(SetText), new object[] { "\tCheck Ping Status", txtGUIFunctionTestInformation });
            while (true)
            {
                if (pingTimer.ElapsedMilliseconds > (Convert.ToInt32(5000)))
                {
                    break;
                }

                if (PingClient(ls_LoginSettingParametersGUI.GatewayIP, 1000))
                {
                    pingStatus = true;
                    TestStatus = "Ping Successfully";
                    Invoke(new SetTextCallBack(SetText), new object[] { "\tPing Successfully!!", txtGUIFunctionTestInformation });
                    break;
                }

                Thread.Sleep(1000);
                string Info = String.Format(".");
                Invoke(new SetTextCallBack(SetText), new object[] { Info, txtGUIFunctionTestInformation });
            }

            if (!pingStatus)
            {
                TestStatus = "Ping Failed";
                Invoke(new SetTextCallBack(SetText), new object[] { "\tPing Failed!!", txtGUIFunctionTestInformation });
                return false;
            }





            return true;
        }

        private void WriteTestReportGUI(bool TestResult, FinalScriptDataGUI FinalScriptPara)
        {

            st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 2] = FinalScriptPara.TestItem;
            st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 2];
            SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "C");

            st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 3] = FinalScriptPara.FunctionName;
            st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 3];
            SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "L");


            st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 4] = FinalScriptPara.Name;
            st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 4];
            SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "L");

            st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 5] = FinalScriptPara.TestResult;
            st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 5];
            SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "C");
            if (FinalScriptPara.TestResult == "PASS")
            {
                st_ExcelObjectGUI.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
            }
            else if (FinalScriptPara.TestResult == "FAIL" || FinalScriptPara.TestResult == "ERROR")
            {
                st_ExcelObjectGUI.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }


            st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 6] = FinalScriptPara.Comment;
            st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 6];
            SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "C");


            st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 7] = FinalScriptPara.Log;
            st_ExcelObjectGUI.excelRange = st_ExcelObjectGUI.excelWorkSheet.Cells[i_DATA_FinalROWGUI, 7];
            //st_ExcelObjectGUI.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            SettingExcelAlignment(st_ExcelObjectGUI.excelRange, "H", "C");
            st_ExcelObjectGUI.excelWorkSheet.Hyperlinks.Add(st_ExcelObjectGUI.excelRange, s_LogFileGUI, Type.Missing, "ConsoleLog", "ConsoleLog");

            SettingExcelFontGUI(st_ExcelObjectGUI.excelRange, "Times New Roman", 11, "B", 3);
            i_DATA_FinalROWGUI++;
        }

        private void WriteWebGUITestReportGUI(int Index, bool TestResult, string ExceptionInfo)
        {
            string Action = st_ReadScriptDataGUI[Index].Action;

            excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 2] = st_ReadScriptDataGUI[Index].TestIndex;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 2];
            SettingExcelAlignment(excelRangeCommon, "H", "C");

            excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 3] = st_ReadScriptDataGUI[Index].ActionName;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 3];
            SettingExcelAlignment(excelRangeCommon, "H", "L");

            excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 4] = Action;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 4];
            SettingExcelAlignment(excelRangeCommon, "H", "C");

            excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 5] = st_ReadScriptDataGUI[Index].ElementType;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 5];
            SettingExcelAlignment(excelRangeCommon, "H", "L");

            excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 10] = ExceptionInfo;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 10];
            SettingExcelAlignment(excelRangeCommon, "H", "L");
            if (ExceptionInfo.CompareTo("Time Out") == 0)
                SettingExcelFontGUI(excelRangeCommon, "Times New Roman", 11, "B", 3);


            if (Action.CompareTo("Set") == 0)
            {
                excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 6] = st_ReadScriptDataGUI[Index].WriteExpectedValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 6];
                SettingExcelAlignment(excelRangeCommon, "H", "C");
            }
            else if (Action.CompareTo("Get") == 0)
            {
                excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 7] = st_ReadScriptDataGUI[Index].WriteExpectedValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 7];
                SettingExcelAlignment(excelRangeCommon, "H", "C");

                excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 8] = st_ReadScriptDataGUI[Index].GetValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 8];
                SettingExcelAlignment(excelRangeCommon, "H", "C");

                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 9];
                if (TestResult == true)
                {
                    excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 9] = "Pass";
                    SettingExcelFontGUI(excelRangeCommon, "Times New Roman", 11, "B", 5);
                }
                else
                {
                    excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 9] = "Fail";
                    SettingExcelFontGUI(excelRangeCommon, "Times New Roman", 11, "B", 3);
                }
                SettingExcelAlignment(excelRangeCommon, "H", "C");
            }

            i_DATA_ROWGUI = i_DATA_ROWGUI + 1;
        }

        private void WriteWebGUITestReportGUI(bool TestResult, CBT_SeleniumApi.GuiScriptParameter ScriptPara)
        {
            string Action = ScriptPara.Action;

            excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 2] = ScriptPara.Index;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 2];
            SettingExcelAlignment(excelRangeCommon, "H", "C");

            excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 3] = ScriptPara.ActionName;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 3];
            SettingExcelAlignment(excelRangeCommon, "H", "L");

            excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 4] = Action;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 4];
            SettingExcelAlignment(excelRangeCommon, "H", "C");

            excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 5] = ScriptPara.ElementType;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 5];
            SettingExcelAlignment(excelRangeCommon, "H", "L");

            excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 10] = ScriptPara.Note;
            excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 10];
            SettingExcelAlignment(excelRangeCommon, "H", "L");
            if (ScriptPara.Note.CompareTo("Time Out") == 0)
                SettingExcelFontGUI(excelRangeCommon, "Times New Roman", 11, "B", 3);


            if (Action.CompareTo("Set") == 0)
            {
                excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 6] = ScriptPara.WriteValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 6];
                SettingExcelAlignment(excelRangeCommon, "H", "C");
            }
            else if (Action.CompareTo("Get") == 0)
            {
                excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 7] = ScriptPara.ExpectedValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 7];
                SettingExcelAlignment(excelRangeCommon, "H", "C");

                excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 8] = ScriptPara.GetValue;
                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 8];
                SettingExcelAlignment(excelRangeCommon, "H", "C");

                excelRangeCommon = excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 9];
                if (TestResult == true)
                {
                    excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 9] = "Pass";
                    SettingExcelFontGUI(excelRangeCommon, "Times New Roman", 11, "B", 5);
                }
                else
                {
                    excelWorkSheetCommon.Cells[i_DATA_ROWGUI, 9] = "Fail";
                    SettingExcelFontGUI(excelRangeCommon, "Times New Roman", 11, "B", 3);
                }
                SettingExcelAlignment(excelRangeCommon, "H", "C");
            }

            i_DATA_ROWGUI = i_DATA_ROWGUI + 1;
        }

        private bool StopWhenTestErrorGUI()
        {
            if (tbs_TestBrowerParametersGUI.StopWhenTestError == true)
                return true;
            return false;
        }

        private void TestFinishedActionGUI()
        {
            SetExcelCellsHighGUI();
            SaveEndTime_and_Version_to_ExcelGUI();
            Save_and_Close_Excel_File_Common(st_ExcelObjectGUI);
            Thread.Sleep(3000);
        }

        private void TestFinishedWebGUIActionGUI()
        {
            SetWebGUIExcelCellsHighGUI();
            SaveEndTime_and_Version_to_WebGUIExcelGUI();
            Save_and_Close_Excel_File();
            
            try
            {
                cs_BrowserGUI.Close_WebDriver();
            }
            catch { }
            Thread.Sleep(3000);
        }

        private bool Re_LoginGUI()
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

        private void SwitchToRGconsoleGUI()
        {
            if (ckboxGUIFunctionTestConsoleLog.Checked == true)
            {
                string ConText = string.Empty;
                string newLine = "\r\n\r\n\r\n";

                cp_ComPortGUI.Write(newLine);
                Thread.Sleep(500);
                ConText += ReadConsoleLogByBytesGUI(cp_ComPortGUI);
                //MessageBox.Show("ConText1:"+ ConText);
                if (ConText.ToLower().IndexOf("cm>") >= 0)
                {
                    cp_ComPortGUI.Write("cd Con\r\n");
                    Thread.Sleep(500);
                }

                ConText = string.Empty;
                cp_ComPortGUI.Write(newLine);
                Thread.Sleep(500);
                ConText += ReadConsoleLogByBytesGUI(cp_ComPortGUI);
                //MessageBox.Show("ConText2:"+ ConText);
                if (ConText.ToLower().IndexOf("cm/cm_console>") >= 0)
                {
                    cp_ComPortGUI.Write("switch\r\n");
                    Thread.Sleep(500);
                }

                ConText = string.Empty;
                cp_ComPortGUI.Write(newLine);
                Thread.Sleep(500);
                ConText += ReadConsoleLogByBytesGUI(cp_ComPortGUI);
                //MessageBox.Show("ConText3:"+ConText);
                if (ConText.ToLower().IndexOf("rg>") >= 0)
                {
                    cp_ComPortGUI.Write("cd Con\r\n");
                    Thread.Sleep(500);
                }
            }
        }

        private bool CheckLastDataGUI(int DataIndex)
        {
            try
            {
                string NextDataTestIndex = st_ReadScriptDataGUI[DataIndex + 1].TestIndex;
                if (st_ReadScriptDataGUI[DataIndex].TestIndex.CompareTo(NextDataTestIndex) != 0)
                    return true;
            }
            catch
            {
                int LastDataIndex = st_ReadScriptDataGUI.Length - 1;
                if (DataIndex == LastDataIndex)
                    return true;
            }
            return false;
        }

        private bool ExceptionActionGUI(string PageURL, ref CBT_SeleniumApi.GuiScriptParameter ScriptPara)
        {
            if (Re_LoginGUI())
            {
                try
                {
                    cs_BrowserGUI.GoToURL(PageURL);
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

        private void loadFinalScriptGUI()
        {
            string filePath = txtGUIFunctionTestFinalScriptFilePath.Text;
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
                if (excelReadRangeCommon[1][itemRow].Text.ToLower() != ("v")|| excelReadRangeCommon[1][itemRow].Text == "" || excelReadRangeCommon[1][itemRow].Text == null)
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
            st_ReadFinalScriptDataGUI = new FinalScriptDataGUI[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadFinalScriptDataGUI[SCRIPT_INDEX].TestItem = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadFinalScriptDataGUI[SCRIPT_INDEX].FunctionName = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadFinalScriptDataGUI[SCRIPT_INDEX].Name = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadFinalScriptDataGUI[SCRIPT_INDEX].StartIndex = ScriptDataArray[SCRIPT_INDEX, 3];
                st_ReadFinalScriptDataGUI[SCRIPT_INDEX].StopIndex = ScriptDataArray[SCRIPT_INDEX, 4];
                st_ReadFinalScriptDataGUI[SCRIPT_INDEX].TestResult = "";
                st_ReadFinalScriptDataGUI[SCRIPT_INDEX].Comment = "";
                st_ReadFinalScriptDataGUI[SCRIPT_INDEX].Log = "";
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
        }

        private void loadStepsScriptGUI()
        {
            string filePath = txtGUIFunctionTestStepsScriptFilePath.Text;
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
            st_ReadStepsScriptDataGUI = new StepsScriptDataGUI[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].Index = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].FunctionName = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].Name = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].Command = new string[20];
                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].WriteValue = new string[20];
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
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].Action = tmpData.Split('$')[1].Trim();
                                break;

                            case "Command":
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].Command[commandlistIndex] = tmpData.Split('$')[1].Trim();
                                commandlistIndex++;
                                break;

                            case "WriteValue":
                                //string[] stmpData = tmpData.Split('$');
                                //int tmpDataLength = stmpData.Length;
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].WriteValue[writeValuelistIndex] = tmpData.Split('$')[tmpData.Split('$').Length - 1].Trim();
                                writeValuelistIndex++;
                                break;

                            case "ExpectedValue":
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].ExpectedValue = tmpData.Split('$')[1].Trim();
                                break;

                            case "TimeOut":
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].TestTimeOut = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpUsername":
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].FtpUsername = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpPassword":
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].FtpPassword = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpIP":
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].FtpIP = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpServerFolder":
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].FtpServerFolder = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpServerPort":
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].FtpServerPort = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpRenameTo":
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].FtpRenameTo = tmpData.Split('$')[1].Trim();
                                break;

                            case "FtpPassiveMode":
                                st_ReadStepsScriptDataGUI[SCRIPT_INDEX].FtpPassiveMode = tmpData.Split('$')[1].Trim();
                                break;
                        }
                    }
                }
            }
            //progressbarForm.Dispose();
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀

        }

        private void loadWebGUIScriptGUI()
        {
            string filePath = txtGUIFunctionTestScriptFilePath.Text;
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
            st_ReadScriptDataGUI = new ScriptDataGUI[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadScriptDataGUI[SCRIPT_INDEX].Procedure = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadScriptDataGUI[SCRIPT_INDEX].TestIndex = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadScriptDataGUI[SCRIPT_INDEX].Action = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadScriptDataGUI[SCRIPT_INDEX].ActionName = ScriptDataArray[SCRIPT_INDEX, 3];
                st_ReadScriptDataGUI[SCRIPT_INDEX].ElementType = ScriptDataArray[SCRIPT_INDEX, 4];
                st_ReadScriptDataGUI[SCRIPT_INDEX].WriteExpectedValue = ScriptDataArray[SCRIPT_INDEX, 5];
                st_ReadScriptDataGUI[SCRIPT_INDEX].RadioButtonExpectedValueXpath = ScriptDataArray[SCRIPT_INDEX, 6];
                st_ReadScriptDataGUI[SCRIPT_INDEX].ElementXpath = ScriptDataArray[SCRIPT_INDEX, 7];
                st_ReadScriptDataGUI[SCRIPT_INDEX].TestTimeOut = ScriptDataArray[SCRIPT_INDEX, 8];
                st_ReadScriptDataGUI[SCRIPT_INDEX].GetValue = "";
            }

            //MessageBox.Show("sdReadScriptData.Length: " + sdReadScriptData.Length);
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal); // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
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
