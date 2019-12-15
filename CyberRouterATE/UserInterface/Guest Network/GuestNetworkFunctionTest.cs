//---------------------------------------------------------------------------------------//
//  This code was created by CyberTan Sally Lee.                                         // 
//  File           : GuestNetworkFunctionTest.cs                                         // 
//  Update         : 2019-03-12                                                          //
//  Version        : 1.0.190312                                                          //
//  Description    :                                                                     //  
//  Modified       : 2019-03-12 Initial version                                          //
//  History        : 2019-03-12 Initial version                                          //  
//                                                                                       //
//---------------------------------------------------------------------------------------//

#define DEBUG_MODE
#undef DEBUG_MODE

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


namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        //********* RouterTestMain.cs ********//
        #region
        //FinalTestItemsScriptDataGuestNetwork[] st_ReadTestItemsScriptGuestNetwork;
        //StepsScriptDataGuestNetwork[] st_ReadStepsScriptDataGuestNetwork;
        //DeviceTestScriptDataGuestNetwork[] st_ReadDeviceTestScriptDataGuestNetwork;

        //int TEST_ITEMS_STRUCT_INDEX = 0;
        //int STEP_SCRIPT_STRUCT_INDEX = 0;
        //int DEVICE_TEST_SCRIPT_STRUCT_INDEX = 0;
        //int i_TestRunWebGuiCtrl = 1;
        #endregion

        //***** WebGuiControlFunction.cs *****//
        #region
        //private delegate void WebGuiCtrlCommonDelegate_SetTextCallBack(string text, TextBox textbox);
        //public delegate void WebGuiCtrlCommonDelegate();

        //Thread threadWebGuiCtrlFT;
        //Thread threadWebGuiCtrlFTstopEvent;
        //Thread threadWebGuiCtrlFTscriptItem;
        //bool bWebGuiCtrlFTThreadRunning = false;
        //bool bWebGuiCtrlTestComplete = true;
        //bool bWebGuiCtrlSingleScriptItemRunning = false;

        //string s_CurrentURL_WebGuiFwUpDnGrade = string.Empty;
        #endregion



        int i_FinalReportStartRow_GuestNetwork = 16;
        string s_ScriptPath_GuestNetwork = System.Windows.Forms.Application.StartupPath + @"\testCondition\GuestNetwork";

        string s_ReportFolderFullPath_GuestNetwork;
        string s_FolderNameGuestNetwork = "GuestNetwork_DATE";
        string s_FinalReportFileName_GuestNetwork = string.Empty;
        string s_FinalReportFilePath_GuestNetwork;
        string s_WebGuiReportFileName_GuestNetwork = string.Empty;
        string s_WebGuiReportFilePath_GuestNetwork;


        ModelStructGuestNetwork struct_ModelStructGuestNetwork = new ModelStructGuestNetwork();
        DeviceSettingGuestNetwork struct_DeviceSettingGuestNetwork = new DeviceSettingGuestNetwork();
        GAsettingGuestNetwork struct_GAsettingGuestNetwork = new GAsettingGuestNetwork();
        TestSettingGuestNetwork struct_TestSettingGuestNetwork = new TestSettingGuestNetwork();

        ExcelObject st_ExcelObjectFinalRepor_GuestNetwork = new ExcelObject();
        ////ExcelObject st_ExcelObjectWebGuiRepor_WebGuiFwUpDnGrade = new ExcelObject();



        //**********************************************************************************//
        //----------------------- Guest Network Function Test Event ------------------------//
        //**********************************************************************************//
#region Guest Network Test Event
        
        private void GuestNetworkTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //----- Hide All TabPage -----//
            ToggleToolStripMenuItem(false);
            tabControl_RouterStartPage.Hide();
            
            //--- Show TabPage(needed) ---//
            tabControl_GuestNetworkTest.Show();
            this.tpGuestNetworkFunctionTest.Parent = this.tabControl_GuestNetworkTest;      //show Function Test TabPag
            
            GuestNetworkTestToolStripMenuItem.Checked = true;
            tsslMessage.Text = tabControl_GuestNetworkTest.TabPages[tabControl_GuestNetworkTest.SelectedIndex].Text + " Control Panel";

            txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Text = System.Windows.Forms.Application.StartupPath + @"\testCondition\DutWebScript\EA7300_WebGuiScript_.xlsx";
            txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Select(txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Text.Length, 0);  // 檔名很長，讓游標位置靠最右邊

            txtGuestNetworkFunctionTest_TestItemsScriptFilePath.Text = System.Windows.Forms.Application.StartupPath + @"\testCondition\GuestNetwork\GuestNetworkFinalTestItems.xlsx";
            txtGuestNetworkFunctionTest_TestItemsScriptFilePath.Select(txtGuestNetworkFunctionTest_TestItemsScriptFilePath.Text.Length, 0);  // 檔名很長，讓游標位置靠最右邊
        }

        private void btnGuestNetworkFunctionTestDeviceTestScriptBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogGuestNetwork = new OpenFileDialog();
            openFileDialogGuestNetwork.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\DutWebScript"; ;

            // Set filter for file extension and default file extension
            openFileDialogGuestNetwork.Filter = "Excel file|*.xlsx";

            // If the file name is not an empty string open it for opening.
            if (openFileDialogGuestNetwork.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialogGuestNetwork.FileName != "")
            {
                txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Clear();
                txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Text = openFileDialogGuestNetwork.FileName;
                txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Select(txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Text.Length, 0); // 檔名很長，讓游標位置靠最右邊
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            }
            else if (txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Text == "" || txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Text == null)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                MessageBox.Show("Please Select a Script File!!");
                return;
            }
        }

        private void btnGuestNetworkFunctionTestSaveInfo_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDebugFileDialog = new SaveFileDialog();

            saveDebugFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            saveDebugFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveDebugFileDialog.FilterIndex = 1;
            saveDebugFileDialog.RestoreDirectory = true;
            saveDebugFileDialog.FileName = "GuestNetworkLog";

            if (saveDebugFileDialog.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText(saveDebugFileDialog.FileName, TEXTBOX_INFO.Text);
            }
        }
        
        private void btnGuestNetworkFunctionTestRun_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Run Btn Checked!!");

            //btnGuestNetworkFunctionTestRun.Enabled = false;
            //System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            TEXTBOX_INFO = txtGuestNetworkFunctionTestInformation;

            if (bWebGuiCtrlFTThreadRunning == false && bWebGuiCtrlTestComplete == true)
            {
                #region

                //-------------- Save Test Parameter --------------//
                SaveModelParameterGuestNetworkFunctionTest();
                SaveDeviceSettingParameterGuestNetworkFunctionTest();
                SaveGAsettingParameterGuestNetworkFunctionTest();
                SaveTestSettingParameterGuestNetworkFunctionTest();



                //-------------------------------------------------//
                //----------------- Start Test --------------------//
                //-------------------------------------------------//

                bWebGuiCtrlFTThreadRunning = true;
                bWebGuiCtrlTestComplete = false;
                ToggleGuestNetworkFunctionTestController(false);  /* Disable all controller */
                TEXTBOX_INFO.Clear();

                //--------------- Start Test Thread ---------------//
                SetText("----------------------------- Start Device GUI ATE Test -----------------------------", TEXTBOX_INFO);
                threadWebGuiCtrlFT = new Thread(new ThreadStart(DoGuestNetworkFunctionTest));
                threadWebGuiCtrlFT.Name = "";
                threadWebGuiCtrlFT.Start();

                //*** Use another thread to catch the stop event of test thread ***//
                threadWebGuiCtrlFTstopEvent = new Thread(new ThreadStart(GuestNetwork_threadWebGuiCtrlFT_StopEvent));
                threadWebGuiCtrlFTstopEvent.Name = "";
                threadWebGuiCtrlFTstopEvent.Start();
                #endregion
            }
            else if (bWebGuiCtrlFTThreadRunning == true && bWebGuiCtrlTestComplete == false)
            {
                tsslMessage.Text = "Function Test Control Panel";
                bWebGuiCtrlFTThreadRunning = false;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "The Test will stop later. Please wait!", "ATE Information", MessageBoxButtons.OK);
            }
        }

#endregion



        //**********************************************************************************//
        //----------------------- Guest Network Function Test Module -----------------------//
        //**********************************************************************************//
#region Guest Network Test Module

        //=================================================================//
        //------------------------- Initial Function ----------------------//
        //=================================================================//
        #region Initial Function
        
        //private void InitWebGuiFwUpDnGradeTestDataView()
        //{
        //    //----- Initial Title -----//
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.ColumnCount = 3;
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.Name = "Test Data View Setting";
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[0].Name = "Test Case";
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[1].Name = "Pass";
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[2].Name = "Fail";

        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[0].Width = 120;
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[1].Width = 70;
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;



        //    //----- Initial Test Status -----//

        //    /** Remove the all data **/
        //    if (dgvWebGuiFwUpDnGradeFunctionTestDataView.RowCount > 1)
        //    {
        //        DataTable dt = (DataTable)dgvWebGuiFwUpDnGradeFunctionTestDataView.DataSource;
        //        dgvWebGuiFwUpDnGradeFunctionTestDataView.Rows.Clear();
        //        dgvWebGuiFwUpDnGradeFunctionTestDataView.DataSource = dt;
        //    }

        //    FwDowngrade.PassCount = 0;
        //    FwDowngrade.FailCount = 0;
        //    FwDowngrade.GridPassRow = 0;    // GridPass(Row,Cell) = (0,1)
        //    FwDowngrade.GridPassCell = 1;
        //    FwDowngrade.GridFailRow = 0;    // GridFail(Row,Cell) = (0,2)
        //    FwDowngrade.GridFailCell = 2;
        //    FwDowngrade.ExcelPassRow = 11;  // ExcelPass(Row,Cell) = (11,9)
        //    FwDowngrade.ExcelPassCell = 9;
        //    FwDowngrade.ExcelFailRow = 11;  // ExcelFail(Row,Cell) = (11,10)
        //    FwDowngrade.ExcelFailCell = 10;

        //    FwUpgrade.PassCount = 0;
        //    FwUpgrade.FailCount = 0;
        //    FwUpgrade.GridPassRow = 1;      // GridPass(Row,Cell) = (1,1)
        //    FwUpgrade.GridPassCell = 1;
        //    FwUpgrade.GridFailRow = 1;      // GridFail(Row,Cell) = (1,2)
        //    FwUpgrade.GridFailCell = 2;
        //    FwUpgrade.ExcelPassRow = 12;    // ExcelPass(Row,Cell) = (12,9)
        //    FwUpgrade.ExcelPassCell = 9;
        //    FwUpgrade.ExcelFailRow = 12;    // ExcelFail(Row,Cell) = (12,10)
        //    FwUpgrade.ExcelFailCell = 10;

        //    TestStatusDict["FW Downgrade"] = FwDowngrade;
        //    TestStatusDict["FW Upgrade"] = FwUpgrade;

        //    AddTeatDataToDataGridView("FW Downgrade");
        //    AddTeatDataToDataGridView("FW Upgrade");

        //}

        //private void AddTeatDataToDataGridView(string TestCase)
        //{
        //    string PassCount = TestStatusDict[TestCase].PassCount.ToString();
        //    string FailCount = TestStatusDict[TestCase].FailCount.ToString();

        //    string[] arrayData = new string[] 
        //    {
        //        TestCase,
        //        PassCount,
        //        FailCount
        //    };
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.Rows.Add(arrayData);
        //}

        //private void InitialTestStatus()
        //{
        //    //----- Initial GUI Data Grid -----//
        //    FwDowngrade.PassCount = 0;
        //    FwDowngrade.FailCount = 0;
        //    FwUpgrade.PassCount = 0;
        //    FwUpgrade.FailCount = 0;

        //    for (int i = 0; i < dgvWebGuiFwUpDnGradeFunctionTestDataView.RowCount - 1; i++)
        //    {
        //        for (int j = 1; j < dgvWebGuiFwUpDnGradeFunctionTestDataView.ColumnCount; j++)
        //        {
        //            Invoke(new AddDataGridViewDelegate_WebGuiFwUpDnGrade(ModifyCurrentTeatStatusToDataGridView), new object[] { i, j, "0" });
        //        }
        //    }

        //    //----- Initial Excel -----//
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[FwDowngrade.ExcelPassRow, FwDowngrade.ExcelPassCell] = "0";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[FwDowngrade.ExcelFailRow, FwDowngrade.ExcelFailCell] = "0";

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[FwUpgrade.ExcelPassRow, FwUpgrade.ExcelPassCell] = "0";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[FwUpgrade.ExcelFailRow, FwUpgrade.ExcelFailCell] = "0";
        //}

        private void TestInitialGuestNetwork()
        {
            RouterTestTimer = null;
            RouterTestTimer = new Stopwatch();
            st_ReadTestItemsScript_RouterTest = null;
            st_ReadStepsScriptData_RouterTest = null;
            st_ReadDeviceWebGuiScriptData_RouterTest = null;

            i_FinalReportStartRow_GuestNetwork = 16;
            s_ScriptPath_GuestNetwork = System.Windows.Forms.Application.StartupPath + @"\testCondition\GuestNetwork";
            s_RouterReportPath = System.Windows.Forms.Application.StartupPath + @"\report";
            s_FolderNameGuestNetwork = "GuestNetwork_DATE";
            s_FinalReportFileName_GuestNetwork = string.Empty;
            s_WebGuiReportFileName_GuestNetwork = string.Empty;

            //InitialTestStatus();

            csbt_TestBrowserType = csbt_TestBrowserChrome;
            if (!WebDriverInitial())
            {
                CloseExcelReportFile(st_ExcelObjectFinalRepor_GuestNetwork);
                ClodeWebDriver();
                bCommonFTThreadRunning = false;
                bWebGuiCtrlSingleScriptItemRunning = false;
                return;
            }
        }

        //private bool WebDriverInitial()
        //{
        //    if (cs_DutFwGuiBrowser == null)
        //    {
        //        cs_DutFwGuiBrowser = new CBT_SeleniumApi();

        //        if (!cs_DutFwGuiBrowser.init(csbt_TestBrowserWebGuiFwUpDnGrade))
        //        {
        //            CloseExcelReportFile();
        //            ClodeWebDriver();
        //            return false;
        //        }
        //        //cs_DutFwGuiBrowser.SettingTimeout(60);
        //        //cs_DutFwGuiBrowser.WindowMaximize();
        //        //Thread.Sleep(1000);
        //        //cs_DutFwGuiBrowser.WindowMinimize();
        //        Thread.Sleep(3000);
        //    }

        //    return true;
        //}

        #endregion


        //=================================================================//
        //----------------------- Main Test Structure ---------------------//
        //=================================================================//
        #region Main Test Structure

        private void GuestNetwork_threadWebGuiCtrlFT_StopEvent()
        {
            while (bWebGuiCtrlFTThreadRunning == true)
            {
                Thread.Sleep(500);
            }

            GuestNetworkFunctionTestThreadFinished();
        }

        private void GuestNetworkFunctionTestThreadFinished()
        {
            Invoke(new WebGuiCtrlCommonDelegate_SetTextCallBack(SetText), new object[] { "------------------------------ Device GUI ATE Test End ------------------------------", TEXTBOX_INFO });
            this.Invoke(new WebGuiCtrlCommonDelegate(ToggleGuestNetworkFunctionTestSettingTrue));
        }

        private void TestFinishedAction_GuestNetwork()
        {
            //CloseExcelReportFile();
            ClodeWebDriver();
            Thread.Sleep(3000);
        }

        //private void CloseExcelReportFile()
        //{
        //    try
        //    {
        //        SaveEndTimetoExcel_Common(st_ExcelObjectFinalRepor_GuestNetwork, 1, 7, 3);
        //        SetExcelCellsHigh_Common(st_ExcelObjectFinalRepor_GuestNetwork, "B16", "B50000");
        //        Save_and_Close_Excel_File_Common(st_ExcelObjectFinalRepor_GuestNetwork);
        //    }
        //    catch { }
        //}

        //private void ClodeWebDriver()
        //{
        //    try
        //    {
        //        bool bResult = cs_DutFwGuiBrowser.Close_WebDriver();
        //        cs_DutFwGuiBrowser = null;
        //    }
        //    catch { }
        //    Thread.Sleep(2000);
        //}

        //private bool CheckWebGuiXPath(ref CBT_SeleniumApi.GuiScriptParameter ScriptPara)
        //{
        //    bool checkXPathResult = false;
        //    string strInfo = string.Empty;
        //    Stopwatch waitingtime = new Stopwatch();

        //    waitingtime.Reset();
        //    waitingtime.Start();
        //    do
        //    {
        //        checkXPathResult = cs_DutFwGuiBrowser.CheckXPathDisplayed(ScriptPara.ElementXpath);
        //        Thread.Sleep(100);

        //    } while (waitingtime.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);
        //    waitingtime.Stop();
        //    Thread.Sleep(2000);

        //    return checkXPathResult;
        //}

        //private bool InTimePeriod()
        //{
        //    string sInfo = string.Empty;
        //    DateTime SettingStartDay = dtpWebGuiFwUpDnGradeDeviceSettingPeriodStartDay.Value;
        //    DateTime SettingEndDay = dtpWebGuiFwUpDnGradeDeviceSettingPeriodEndDay.Value;

        //    //--- 省掉日期資訊，只關心時間 ---//
        //    string SystemTimeNow = DateTime.Now.ToString("HH:mm:ss");
        //    string SettingStartTime = SettingStartDay.ToString("HH:mm:ss");
        //    string SettingEndTime = SettingEndDay.ToString("HH:mm:ss");

        //    DateTime dtStartTime = Convert.ToDateTime(SettingStartTime);
        //    DateTime dtEndTime = Convert.ToDateTime(SettingEndTime);
        //    DateTime dtTimeNow = Convert.ToDateTime(SystemTimeNow);

        //    //int i = DateTime.Compare(dt1, dt3); // dt1 比 dt3大 --> 1
        //    //int k = DateTime.Compare(dt2, dt3); // dt2 比 dt3小 --> -1

        //    if (DateTime.Compare(dtTimeNow, dtStartTime) >= 0 && DateTime.Compare(dtTimeNow, dtEndTime) <= 0)
        //        return true;

        //    return false;
        //}
        
        //private void SaveWebGuiTestResult(ref CBT_SeleniumApi.GuiScriptParameter ScriptPara, bool bTestResult)
        //{
        //    if (ScriptPara.TestResult == null || ScriptPara.TestResult == "")
        //    {
        //        if (bTestResult == true || ScriptPara.Note == string.Empty || ScriptPara.Note == "")
        //            ScriptPara.TestResult = "PASS";
        //        else
        //            ScriptPara.TestResult = "FAIL";
        //    }

        //    //***** WriteStepTestResultToArray *****//
        //    //-- st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest]   <-- Current Step
        //    string CuurentStepTestResult = st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].TestResult;
        //    if (ScriptPara.TestResult.CompareTo("FAIL") == 0)
        //    {
        //        st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].TestResult = "FAIL";
        //        st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Note = ScriptPara.Note;
        //    }
        //}

        private void DoGuestNetworkFunctionTest()
        {
            //MessageBox.Show("DoGuestNetworkFunctionTest()");

            string sInfo = string.Empty;
            int TestTimeOut = 10000;
            int iTestRun = Convert.ToInt32(struct_TestSettingGuestNetwork.TestRun);

            this.Invoke(new WebGuiCtrlCommonDelegate(ToggleGuestNetworkFunctionTestSystemCursorWait));


            //    CreateNewFinalReport_WebGuiFwUpDnGrade();
            //    //CreateNewWebGuiReport_WebGuiFwUpDnGrade();
            TestInitialGuestNetwork();

            
            
            sInfo = string.Format("Loading Test Items Script... ");
            //Invoke(new SetTextCallBack(SetText), new object[] { sInfo, txtGuestNetworkFunctionTestInformation });
            Invoke(new SetTextCallBack(SetText), new object[] { sInfo, TEXTBOX_INFO });
            GuestNetworkFunctionTestLoadFinalTestItemsScript();

            sInfo = string.Format("Loading Steps Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { sInfo, TEXTBOX_INFO });
            GuestNetworkFunctionTestLoadStepsScript();

            sInfo = string.Format("Loading WebGUI Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { sInfo, TEXTBOX_INFO });
            GuestNetworkFunctionTestLoadWebGUIScript();



            for (i_TestRun_RouterTest = 1; i_TestRun_RouterTest <= iTestRun; i_TestRun_RouterTest++)
            {
                sInfo = string.Format("****************** [ Test Times: {0} ] ******************", i_TestRun_RouterTest.ToString());
                Invoke(new SetTextCallBack(SetText), new object[] { "", TEXTBOX_INFO });
                Invoke(new SetTextCallBack(SetText), new object[] { sInfo, TEXTBOX_INFO });
                this.Invoke(new WebGuiCtrlCommonDelegate(ToggleGuestNetworkFunctionTestSystemCursorDefault));

                for (TEST_ITEMS_SCRIPT_ST_IDX_RouterTest = 0; TEST_ITEMS_SCRIPT_ST_IDX_RouterTest < st_ReadTestItemsScript_RouterTest.Length; TEST_ITEMS_SCRIPT_ST_IDX_RouterTest++)                                            //------------------ Final Script Array-----------------//
                {
                    #region Final Script Array

                    if (!WebDriverInitial())   //------------- 每個測項結束後完都會 Close Web Driver, 然後重新開啟
                    {
                        CloseExcelReportFile(st_ExcelObjectFinalRepor_GuestNetwork);
                        ClodeWebDriver();
                        sInfo = string.Format("### FAIL in :  <Test Item> {0}  \n### Note: WebDriver Initial() FAIL!!", st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].TestItem);
                        st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].TestResult = "FAIL";
                        st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].Comment = sInfo;
                        Invoke(new SetTextCallBack(SetText), new object[] { "", TEXTBOX_INFO });
                        Invoke(new SetTextCallBack(SetText), new object[] { sInfo, TEXTBOX_INFO });
                        //WriteFinalTestResultExcel(st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest]); //////////////////////
                        continue;
                    }

                    int iStepsStartIndex = Convert.ToInt32(st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].StartIndex);
                    int iStepsStopIndex = Convert.ToInt32(st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].StopIndex);
                    sInfo = string.Format("***** <Test Item> : {0} *****", st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].TestItem);
                    Invoke(new SetTextCallBack(SetText), new object[] { "", TEXTBOX_INFO });
                    Invoke(new SetTextCallBack(SetText), new object[] { sInfo, TEXTBOX_INFO });

                    bool bSkipTestStep = false;

                    for (STEP_SCRIPT_ST_IDX_RouterTest = 0; STEP_SCRIPT_ST_IDX_RouterTest < st_ReadStepsScriptData_RouterTest.Length; STEP_SCRIPT_ST_IDX_RouterTest++)                                     //------------------ Step Script Array------------------//
                    {
                        #region Step Script Array

                        int icurStepsIndex = Convert.ToInt32(st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Index);
                        string FunctionName = st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].FunctionName;

                        if (bSkipTestStep == false && icurStepsIndex >= iStepsStartIndex && icurStepsIndex <= iStepsStopIndex)  // iStepsStartIndex <=  icurStepsIndex <= iStepsStopIndex
                        {
                            string sStepName = st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Name;

                            sInfo = string.Format("***    <Test Steps>: {0}", sStepName);
                            Invoke(new SetTextCallBack(SetText), new object[] { "", TEXTBOX_INFO });
                            Invoke(new SetTextCallBack(SetText), new object[] { sInfo, TEXTBOX_INFO });


                            switch (FunctionName)
                            {
                                case "WebGUI":
                                    bSkipTestStep = GuestNetworkFunctionTest_WebGuiCtrlTestLoop(TestTimeOut);
                                    break;
                                case "REMOTE":
                                    bSkipTestStep = GuestNetworkFunctionTest_RemoteTestLoop(TestTimeOut);
                                    break;
                            }



                            #region Old Web GUI code
                            /*
                            for (DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest = 0; DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest < st_ReadDeviceWebGuiScriptData_RouterTest.Length; DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest++)   //------------ Deivce Test Script Data Array -----------//
                            {
                                #region Deivce Test Script Data Array

                                string strCurrentStepAction = st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Action;

                                if (strCurrentStepAction == st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].Procedure)
                                {
                                    #region Device Test Script Main Test Thread

                                    if (st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].TestTimeOut == null || st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].TestTimeOut == "")
                                    {
                                        st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].TestTimeOut = struct_TestSettingGuestNetwork.DefaultTimeout;
                                        TestTimeOut = Convert.ToInt32(struct_TestSettingGuestNetwork.DefaultTimeout) * 1000;
                                    }
                                    else
                                    {
                                        TestTimeOut = Convert.ToInt32(st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].TestTimeOut) * 1000;
                                    }


                                    //----------------------------------------------------------------//
                                    //----------------------- Main Test Thread -----------------------//
                                    //----------------------------------------------------------------//              
                                    RouterTestTimer.Reset();
                                    RouterTestTimer.Start();
                                    bWebGuiCtrlSingleScriptItemRunning = true;
                                    threadWebGuiCtrlFTscriptItem = new Thread(new ThreadStart(GuestNetworkFunctionTest_WebGuiCtrlMainFunction));
                                    //threadWebGuiCtrlFTscriptItem.Name = "";
                                    threadWebGuiCtrlFTscriptItem.Start();

                                    while (RouterTestTimer.ElapsedMilliseconds <= TestTimeOut && bWebGuiCtrlSingleScriptItemRunning == true)
                                    {
                                        Thread.Sleep(100);
                                    }
                                    RouterTestTimer.Stop();


                                    //----------------------//
                                    //------ Time Out ------//
                                    //----------------------//
                                    if (RouterTestTimer.ElapsedMilliseconds > TestTimeOut)
                                    {
                                        st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Note = "Time Out!!   " + st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Note;
                                        bSkipTestStep = true;
                                        break;
                                    }
                                    //----------------------//
                                    //-- Test Result Fail --//
                                    //----------------------//
                                    else if (st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].TestResult == "FAIL")
                                    {
                                        bSkipTestStep = true;
                                        break;
                                    }

                                    //WriteWebGuiTestResultToReport();        /////////////////////////////////// Write Web GUI Test Report

                                    #endregion //Device Test Script Main Test Thread
                                }

                                #endregion //Deivce Test Script Data Array
                            }
                            */
                            #endregion



                            //--- Test Step Exception or Test Result Fail, Save Comment ---//
                            if (bSkipTestStep == true)
                            {
                                GuestNetworkExceptionAction();
                            }

                        }

                        #endregion //Step Script Array
                    }

                    //ClodeWebDriver();  //------------- FW Stress 每次 Up/Downgrade 完都要 Close Web Driver
                    //WriteFinalTestResultToReport();        /////////////////////////////////// Write Final Test Result

                    #endregion //Final Script Array
                }
            }

            bWebGuiCtrlFTThreadRunning = false;   //-- Function Test Thread Running Finished
        }

        private bool GuestNetworkFunctionTest_WebGuiCtrlTestLoop(int TestTimeOut)
        {
            for (DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest = 0; DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest < st_ReadDeviceWebGuiScriptData_RouterTest.Length; DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest++)   //------------ Deivce Test Script Data Array -----------//
            {
                string strCurrentStepAction = st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Action;

                if (strCurrentStepAction == st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].Procedure)
                {
                    if (st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].TestTimeOut == null || st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].TestTimeOut == "")
                    {
                        st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].TestTimeOut = struct_TestSettingGuestNetwork.DefaultTimeout;
                        TestTimeOut = Convert.ToInt32(struct_TestSettingGuestNetwork.DefaultTimeout) * 1000;
                    }
                    else
                    {
                        TestTimeOut = Convert.ToInt32(st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].TestTimeOut) * 1000;
                    }


                    //----------------------------------------------------------------//
                    //----------------------- Main Test Thread -----------------------//
                    //----------------------------------------------------------------//              
                    RouterTestTimer.Reset();
                    RouterTestTimer.Start();
                    bWebGuiCtrlSingleScriptItemRunning = true;
                    threadWebGuiCtrlFTscriptItem = new Thread(new ThreadStart(GuestNetworkFunctionTest_WebGuiCtrlMainFunction));
                    //threadWebGuiCtrlFTscriptItem.Name = "";
                    threadWebGuiCtrlFTscriptItem.Start();

                    while (RouterTestTimer.ElapsedMilliseconds <= TestTimeOut && bWebGuiCtrlSingleScriptItemRunning == true)
                    {
                        Thread.Sleep(100);
                    }
                    RouterTestTimer.Stop();


                    //----------------------//
                    //------ Time Out ------//
                    //----------------------//
                    if (RouterTestTimer.ElapsedMilliseconds > TestTimeOut)
                    {
                        st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Note = "Time Out!!   " + st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Note;
                        return true;
                    }
                    //----------------------//
                    //-- Test Result Fail --//
                    //----------------------//
                    else if (st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].TestResult == "FAIL")
                    {
                        return true;
                    }

                    //WriteWebGuiTestResultToReport();        /////////////////////////////////// Write Web GUI Test Report
                }
            }

            return false;
        }

        private bool GuestNetworkFunctionTest_RemoteTestLoop(int TestTimeOut)
        {
            TestTimeOut = Convert.ToInt32(struct_TestSettingGuestNetwork.DefaultTimeout) * 1000;
            string RemoteStr = st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].RemoteSendString;



            return false;
        }

        private void GuestNetworkFunctionTest_WebGuiCtrlMainFunction()
        {
            //MessageBox.Show("GuestNetworkFunctionTestMainFunction()");

            Thread.Sleep(2000);
            CBT_SeleniumApi.GuiScriptParameter ScriptPara = new CBT_SeleniumApi.GuiScriptParameter();
            string sLoginURL = string.Format("http://{0}:{1}", struct_DeviceSettingGuestNetwork.DeviceIP, struct_DeviceSettingGuestNetwork.HTTP_Port);
            string sExceptionInfo = string.Empty;
            string strInfo = string.Empty;
            bool bTestResult = false;
            bool checkXPathResult = false;

            ScriptPara.Procedure = st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].Procedure;
            ScriptPara.Index = st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].TestIndex;
            ScriptPara.Action = st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].Action;
            ScriptPara.ActionName = st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].ActionName;

            ScriptPara.ElementType = st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].ElementType;
            ScriptPara.ElementXpath = st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].ElementXpath;
            ScriptPara.ElementXpath = ScriptPara.ElementXpath.Replace('\"', '\'');
            ScriptPara.RadioBtnExpectedValueXpath = st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].RadioButtonExpectedValueXpath.Replace('\"', '\'');
            ScriptPara.WriteValue = st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].WriteExpectedValue;
            ScriptPara.ExpectedValue = st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].WriteExpectedValue;
            ScriptPara.URL = sLoginURL + st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].WriteExpectedValue;
            ScriptPara.TestTimeOut = st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].TestTimeOut;

            strInfo = string.Format("**      [WebGUI Index]: {0}    [Action]:{1}    [Action Name]:{2}    ", ScriptPara.Index, ScriptPara.Action, ScriptPara.ActionName);
            Invoke(new SetTextCallBack(SetText), new object[] { strInfo, TEXTBOX_INFO });
            

            strInfo = string.Empty;
            ScriptPara.Note = string.Empty;
            ScriptPara.GetValue = string.Empty;

            //-----------------------------//
            //----- Login to Main Page-----//
            //-----------------------------//
            //if (DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest == 0)
            //{
            //    s_CurrentURL_WebGuiCtrl = sLoginURL;
            //    bTestResult = cs_DutWebGuiCtrlClass.GoToURL(sLoginURL);
            //    Thread.Sleep(3000);
            //}

            switch (ScriptPara.Action)
            {
                case "Goto":
                    #region Goto URL
                    if (s_CurrentURL_WebGuiCtrl.CompareTo(ScriptPara.URL) != 0)
                    {
                        s_CurrentURL_WebGuiCtrl = ScriptPara.URL;
                        bTestResult = cs_DutWebGuiCtrlClass.GoToURL(s_CurrentURL_WebGuiCtrl);
                        Thread.Sleep(3000);

                        if (bTestResult == false)
                            ScriptPara.Note = string.Format("-----> Goto URL Error!\n({0})", ScriptPara.URL);
                    }
                    break;
                    #endregion

                case "Set":
                    #region Set Value
                    RouterTestTimer.Stop();
                    checkXPathResult = CheckWebGuiXPath(ref ScriptPara);
                    RouterTestTimer.Start();

                    if (checkXPathResult == true)
                    {
                        bTestResult = WebGuiCtrlAction_Set(ref ScriptPara);
                    }
                    else
                    {
                        ScriptPara.Note = string.Format("-----> Couldn't find the Web GUI element!");
                        bTestResult = false;
                    }
                    break;
                    #endregion

                case "Get":
                    #region Get Value
                    RouterTestTimer.Stop();
                    checkXPathResult = CheckWebGuiXPath(ref ScriptPara);
                    RouterTestTimer.Start();

                    if (checkXPathResult == true)
                    {
                        if (ScriptPara.ElementType.CompareTo("RADIO_BUTTON") == 0)
                        {
                            ScriptPara.ElementXpath = ScriptPara.RadioBtnExpectedValueXpath;
                        }
                        bTestResult = WebGuiCtrlAction_Get(ref ScriptPara);
                    }
                    else
                    {
                        ScriptPara.Note = string.Format("-----> Couldn't find the Web GUI element!");
                        bTestResult = false;
                    }
                    break;
                    #endregion

                case "Wait":
                    #region Wait
                    //--------------------//
                    //----- Wait Time-----//
                    //--------------------//
                    if (ScriptPara.ElementXpath.CompareTo("") == 0)
                    {
                        strInfo = string.Format("Waiting for {0} seconds...", ScriptPara.TestTimeOut);
                        Invoke(new SetTextCallBack(SetText), new object[] { strInfo, TEXTBOX_INFO });

                        RouterTestTimer.Stop();
                        Thread.Sleep(Convert.ToInt32(ScriptPara.TestTimeOut) * 1000);
                        RouterTestTimer.Start();
                        bTestResult = true;
                    }
                    
                    //--------------------//
                    //--- Wait Element ---//
                    //--------------------//
                    else
                    {
                        RouterTestTimer.Stop();
                        checkXPathResult = CheckWebGuiXPath(ref ScriptPara);
                        RouterTestTimer.Start();

                        if (checkXPathResult == false)
                        {
                            ScriptPara.Note = string.Format("-----> Couldn't find the Web GUI element!");
                            bTestResult = false;
                        }
                    }
                    break;
                    #endregion

                case "CloseWeb":
                    #region Close Web
                    bTestResult = ClodeWebDriver();
                    if (bTestResult == false)
                    {
                        ScriptPara.Note = string.Format("-----> Close Web Error!");
                    }
                    break;
                    #endregion
            }

            strInfo = ScriptPara.Note;
            if (bTestResult == false)
                Invoke(new SetTextCallBack(SetText), new object[] { strInfo, TEXTBOX_INFO });


            //----- Save Current Test Result -----//
            //////////////////////////////////////////   SaveWebGuiTestResult(ref ScriptPara, bTestResult);

            bWebGuiCtrlSingleScriptItemRunning = false;
        }

        private void GuestNetworkExceptionAction()
        {
            string failInfo = string.Format("### FAIL in: <Test Item> {0} (item index:{1}),   <Test Steps> {2} (step index:{3}),   <Action Name> {4}", st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].TestItem, st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].ItemIndex, st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Name, st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Index, st_ReadDeviceWebGuiScriptData_RouterTest[DEVICE_WEBGUI_SCRIPT_ST_IDX_RouterTest].ActionName);
            Invoke(new SetTextCallBack(SetText), new object[] { "", TEXTBOX_INFO });
            Invoke(new SetTextCallBack(SetText), new object[] { failInfo, TEXTBOX_INFO });

            string failNote = string.Format("### Note: {0}", st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Note);
            Invoke(new SetTextCallBack(SetText), new object[] { failNote, TEXTBOX_INFO });

            //------ Save FAIL Comment ------//
            st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].Comment = failInfo + "\n" + failNote;


            //------ ScreenShot ------//
            string TestItem = st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].TestItem;
            string RunTimes = i_TestRun_RouterTest.ToString();
            string DateTimeStr = DateTime.Now.ToString("yyyyMMdd-HHmmss");

            string ScreenShotFolderPath = s_ReportFolderFullPath_GuestNetwork + @"\ScreenShot";
            bool isFolderExists = System.IO.Directory.Exists(ScreenShotFolderPath);
            if (!isFolderExists)
                System.IO.Directory.CreateDirectory(ScreenShotFolderPath);

            string ScreenShotFilePath = string.Format(ScreenShotFolderPath + @"\{0}_Run{1}_{2}.jpg", TestItem, RunTimes, DateTimeStr);
            PrintFullScreen_Common(ScreenShotFilePath);
            st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].ScreenShotPath = ScreenShotFilePath;
        }

        #endregion


        //=================================================================//
        //----------------------- Save Test Parameter ---------------------//
        //=================================================================//
        #region Save Test Parameter

        private void SaveModelParameterGuestNetworkFunctionTest()
        {
            struct_ModelStructGuestNetwork.ModelName = txtGuestNetworkFunctionTestModelName.Text;
            struct_ModelStructGuestNetwork.SN = txtGuestNetworkFunctionTestModelSerialNumber.Text;
            struct_ModelStructGuestNetwork.SWver = txtGuestNetworkFunctionTestModelSwVersion.Text;
            struct_ModelStructGuestNetwork.HWver = txtGuestNetworkFunctionTestModelHwVersion.Text;
        }

        private void SaveDeviceSettingParameterGuestNetworkFunctionTest()
        {
            struct_DeviceSettingGuestNetwork.DeviceIP = masktxtGuestNetworkFunctionTestDeviceIP.Text;
            struct_DeviceSettingGuestNetwork.HTTP_Port = nudGuestNetworkFunctionTestDefaultSettingsHTTPport.Value.ToString();
            struct_DeviceSettingGuestNetwork.RebootWaitTime = nudGuestNetworkFunctionTestRebootWaitTime.Value.ToString();

            if (ckboxGuestNetworkFunctionTestConsoleLog.Checked)
                struct_DeviceSettingGuestNetwork.ConsoleLog = true;
            else
                struct_DeviceSettingGuestNetwork.ConsoleLog = false;
        }

        private void SaveGAsettingParameterGuestNetworkFunctionTest()
        {
            struct_GAsettingGuestNetwork.RemoteServerIP = masktxtGuestNetworkGAsettingRemoteServerIP.Text;
            struct_GAsettingGuestNetwork.RemoteServerPort = nudGuestNetworkGAsettingRemoteServerPort.Value.ToString();
        }

        private void SaveTestSettingParameterGuestNetworkFunctionTest()
        {
            struct_TestSettingGuestNetwork.DeivceTestScriptPath = txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Text;
            struct_TestSettingGuestNetwork.TestRun = nudGuestNetworkFunctionTestTestRun.Value.ToString();
            struct_TestSettingGuestNetwork.DefaultTimeout = nudGuestNetworkFunctionTestDefaultTimeout.Value.ToString();

            if (ckboxGuestNetworkFunctionTestStopWhenTestError.Checked)
                struct_TestSettingGuestNetwork.StopWhenTestError = true;
            else
                struct_TestSettingGuestNetwork.StopWhenTestError = false;
        }

        #endregion


        //=================================================================//
        //------------------------ Load Test Script -----------------------//
        //=================================================================//
        #region Load Test Script

        private void GuestNetworkFunctionTestLoadFinalTestItemsScript()
        {
            string filePath = txtGuestNetworkFunctionTest_TestItemsScriptFilePath.Text;
            File.SetAttributes(filePath, FileAttributes.ReadOnly);                 // 屬性設定成唯讀
            excelReadAppCommon = new Excel.Application();
            excelReadWorkBookCommon = excelReadAppCommon.Workbooks.Open(filePath); //開啟舊檔案
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
                //如果第一欄沒打勾或任何標記將不讀取
                if (excelReadRangeCommon[1][itemRow].Text == "" || excelReadRangeCommon[1][itemRow].Text == null)
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
            st_ReadTestItemsScript_RouterTest = new FinalTestItemsScriptData_RouterTest[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadTestItemsScript_RouterTest[SCRIPT_INDEX].ItemIndex = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadTestItemsScript_RouterTest[SCRIPT_INDEX].TestItem = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadTestItemsScript_RouterTest[SCRIPT_INDEX].ItemSection = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadTestItemsScript_RouterTest[SCRIPT_INDEX].StartIndex = ScriptDataArray[SCRIPT_INDEX, 3];
                st_ReadTestItemsScript_RouterTest[SCRIPT_INDEX].StopIndex = ScriptDataArray[SCRIPT_INDEX, 4];
                st_ReadTestItemsScript_RouterTest[SCRIPT_INDEX].TestResult = "";
                st_ReadTestItemsScript_RouterTest[SCRIPT_INDEX].Comment = "";
                st_ReadTestItemsScript_RouterTest[SCRIPT_INDEX].Log = "";
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal);   // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
        }

        private void GuestNetworkFunctionTestLoadStepsScript()
        {
            string filePath = System.Windows.Forms.Application.StartupPath + @"\testCondition\GuestNetwork\GuestNetworkSteps.xlsx";
            File.SetAttributes(filePath, FileAttributes.ReadOnly);                  // 屬性設定成唯讀
            excelReadAppCommon = new Excel.Application();
            excelReadWorkBookCommon = excelReadAppCommon.Workbooks.Open(filePath);  //開啟舊檔案
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
            st_ReadStepsScriptData_RouterTest = new StepsScriptData_RouterTest[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadStepsScriptData_RouterTest[SCRIPT_INDEX].Index = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadStepsScriptData_RouterTest[SCRIPT_INDEX].FunctionName = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadStepsScriptData_RouterTest[SCRIPT_INDEX].Name = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadStepsScriptData_RouterTest[SCRIPT_INDEX].Command = new string[20];

                int commandlistIndex = 0;
                for (int DC = 3; DC < ColumnCount; DC++)
                {
                    if (ScriptDataArray[SCRIPT_INDEX, DC] != null && ScriptDataArray[SCRIPT_INDEX, DC] != "")
                    {
                        string tmpData = ScriptDataArray[SCRIPT_INDEX, DC].Replace("::", "$");
                        switch (tmpData.Split('$')[0].Trim())
                        {
                            case "Action":
                                st_ReadStepsScriptData_RouterTest[SCRIPT_INDEX].Action = tmpData.Split('$')[1].Trim();
                                break;

                            case "Command":
                                st_ReadStepsScriptData_RouterTest[SCRIPT_INDEX].Command[commandlistIndex] = tmpData.Split('$')[1].Trim();
                                commandlistIndex++;
                                break;

                            case "ExpectedValue":
                                st_ReadStepsScriptData_RouterTest[SCRIPT_INDEX].ExpectedValue = tmpData.Split('$')[1].Trim();
                                break;

                            case "RemoteSendString":
                                st_ReadStepsScriptData_RouterTest[SCRIPT_INDEX].RemoteSendString = tmpData.Split('$')[1].Trim();
                                break;

                            case "TimeOut":
                                st_ReadStepsScriptData_RouterTest[SCRIPT_INDEX].TestTimeOut = tmpData.Split('$')[1].Trim();
                                break;
                        }
                    }
                }
            }
            //progressbarForm.Dispose();
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal);    // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
        }

        private void GuestNetworkFunctionTestLoadWebGUIScript()
        {
            string filePath = txtGuestNetworkFunctionTestDeviceTestScriptFilePath.Text;
            File.SetAttributes(filePath, FileAttributes.ReadOnly);                  // 屬性設定成唯讀
            excelReadAppCommon = new Excel.Application();
            excelReadWorkBookCommon = excelReadAppCommon.Workbooks.Open(filePath);  //開啟舊檔案
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
            st_ReadDeviceWebGuiScriptData_RouterTest = new DeviceWebGuiScriptData_RouterTest[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadDeviceWebGuiScriptData_RouterTest[SCRIPT_INDEX].Procedure = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadDeviceWebGuiScriptData_RouterTest[SCRIPT_INDEX].TestIndex = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadDeviceWebGuiScriptData_RouterTest[SCRIPT_INDEX].Action = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadDeviceWebGuiScriptData_RouterTest[SCRIPT_INDEX].ActionName = ScriptDataArray[SCRIPT_INDEX, 3];
                st_ReadDeviceWebGuiScriptData_RouterTest[SCRIPT_INDEX].ElementType = ScriptDataArray[SCRIPT_INDEX, 4];
                st_ReadDeviceWebGuiScriptData_RouterTest[SCRIPT_INDEX].WriteExpectedValue = ScriptDataArray[SCRIPT_INDEX, 5];
                st_ReadDeviceWebGuiScriptData_RouterTest[SCRIPT_INDEX].RadioButtonExpectedValueXpath = ScriptDataArray[SCRIPT_INDEX, 6];
                st_ReadDeviceWebGuiScriptData_RouterTest[SCRIPT_INDEX].ElementXpath = ScriptDataArray[SCRIPT_INDEX, 7];
                st_ReadDeviceWebGuiScriptData_RouterTest[SCRIPT_INDEX].TestTimeOut = ScriptDataArray[SCRIPT_INDEX, 8];
                st_ReadDeviceWebGuiScriptData_RouterTest[SCRIPT_INDEX].GetValue = "";
            }

            //MessageBox.Show("sdReadScriptData.Length: " + sdReadScriptData.Length);
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal);   // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
        }

        #endregion


        //=================================================================//
        //----------------------- Report & Excel File ---------------------//
        //=================================================================//
        #region Report & Exce lFile

        //private void CreateNewFinalReport_WebGuiFwUpDnGrade()
        //{
        //    s_FinalReportFileName_GuestNetwork = "MODEL_FinalReport_SN_SW_HW_DATE.xlsx";
        //    s_FinalReportFileName_GuestNetwork = getReportFileNameWebGuiFwUpDnGrade(s_FinalReportFileName_GuestNetwork);

        //    string sSubPath;
        //    bool bIsExists;

        //    s_FolderNameGuestNetwork = s_FolderNameGuestNetwork.Replace("DATE", DateTime.Now.ToString("yyyyMMdd-HHmmss"));
        //    sSubPath = System.Windows.Forms.Application.StartupPath + "\\report\\" + s_FolderNameGuestNetwork;
        //    s_ReportFolderFullPath_GuestNetwork = sSubPath;
        //    bIsExists = System.IO.Directory.Exists(sSubPath);

        //    if (!bIsExists)
        //        System.IO.Directory.CreateDirectory(sSubPath);

        //    s_FinalReportFilePath_GuestNetwork = s_RouterReportPath + @"\" + s_FolderNameGuestNetwork + @"\" + s_FinalReportFileName_GuestNetwork;
        //    initialFinalReportExcel_WebGuiFwUpDnGrade(s_FinalReportFilePath_GuestNetwork);
        //}

        //private void CreateNewWebGuiReport_WebGuiFwUpDnGrade()
        //{
        //    s_WebGuiReportFileName_GuestNetwork = "MODEL_WebGuiReport_SN_SW_HW_DATE.xlsx";
        //    s_WebGuiReportFileName_GuestNetwork = getReportFileNameWebGuiFwUpDnGrade(s_WebGuiReportFileName_GuestNetwork);

        //    s_WebGuiReportFilePath_GuestNetwork = s_RouterReportPath + @"\" + s_FolderNameGuestNetwork + @"\" + s_WebGuiReportFileName_GuestNetwork;
        //    initialWebGuiReportExcel_WebGuiFwUpDnGrade(s_WebGuiReportFilePath_GuestNetwork);
        //}

        //public string getReportFileNameWebGuiFwUpDnGrade(string reportTargetfileName)
        //{
        //    reportTargetfileName = reportTargetfileName.Replace("MODEL", (struct_ModelStructWebGuiFwUpDnGrade.ModelName == "") ? "FwStress" : struct_ModelStructWebGuiFwUpDnGrade.ModelName);
        //    reportTargetfileName = reportTargetfileName.Replace("SN", struct_ModelStructWebGuiFwUpDnGrade.SN);
        //    reportTargetfileName = reportTargetfileName.Replace("SW", struct_ModelStructWebGuiFwUpDnGrade.SWver);
        //    reportTargetfileName = reportTargetfileName.Replace("HW", struct_ModelStructWebGuiFwUpDnGrade.HWver);
        //    reportTargetfileName = reportTargetfileName.Replace("DATE", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
        //    return reportTargetfileName;
        //}

        //private void initialFinalReportExcel_WebGuiFwUpDnGrade(string filePath)
        //{
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp = new Excel.Application();

        //    /*** Set Excel visible ***/
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Visible = true;

        //    /*** Do not show alert ***/
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.DisplayAlerts = false;

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.UserControl = true;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Interactive = false;

        //    //Set font and font size attributes
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.StandardFont = "Times New Roman";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.StandardFontSize = 11;

        //    /*** This method is used to open an Excel workbook by passing the file path as a parameter to this method. ***/
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkBook = st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Workbooks.Add(misValue);

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkBook.Sheets[1];
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Name = "Test Results";
        //    createExcelTitle_WebGuiFwUpDnGrade();

        //    SaveAsNewExcel_Common(st_ExcelObjectFinalRepor_GuestNetwork, filePath);
        //}

        //private void initialWebGuiReportExcel_WebGuiFwUpDnGrade(string filePath)
        //{
        //}

        //private void createExcelTitle_WebGuiFwUpDnGrade()
        //{
        //    //*** SettingExcelFont() ***//
        //    //--- [ColorIndex] 黑:1, 紅:3, 藍:5,淡黃:27, 淡橘:44 
        //    //--- [Font Style] U:底線 B:粗體 I:斜體 (可同時設定多個屬性)   ex:"UBI"、"BI"、"I"、"B"...

        //    //*** SettingExcelAlignment() ***//
        //    //--- [direction] H:水平方向 V:垂直方向 (可同時設定多個屬性)   ex:"HV"、"H"、"V"
        //    //--- [AlignmentType] C:置中 L:靠左 R:靠右


        //    //--------------------------------------- Report Title ---------------------------------------------//
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 2];
        //    SettingExcelFontUSBStorage(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "Times New Roman", 14, "U", 1);
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 2] = "CyberATE Web GUI FW Stress Final Report";

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[4, 2] = "Test";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[5, 2] = "Station";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[6, 2] = "Start Time";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[7, 2] = "End Time";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[8, 2] = "Model";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[9, 2] = "Serial";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[10, 2] = "SW Version";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[11, 2] = "HW Version";

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[4, 3] = "Web GUI FW Stress Test";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[5, 3] = "CyberTAN ATE-" + struct_ModelStructWebGuiFwUpDnGrade.ModelName;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[6, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[8, 3] = struct_ModelStructWebGuiFwUpDnGrade.ModelName;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[9, 3] = struct_ModelStructWebGuiFwUpDnGrade.SN;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[10, 3] = struct_ModelStructWebGuiFwUpDnGrade.SWver;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[11, 3] = struct_ModelStructWebGuiFwUpDnGrade.HWver;
        //    //--------------------------------------------------------------------------------------------------//



        //    //---------------------------------------- Script Info ---------------------------------------------//
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[3, 8], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[7, 9]].Borders.LineStyle = 1;  // 畫表格框線
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[3, 8] = "Option";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[3, 9] = "Value";

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[3, 8], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[3, 9]].Font.Underline = true;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[3, 8], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[3, 9]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[3, 8], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[3, 9]].Font.FontStyle = "Bold";

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range["I4"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[4, 8] = "Test Times";
        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[4, 9] = i_TestRun_RouterTest;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[5, 8] = "Final Script";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[5, 9] = s_ScriptPath_GuestNetwork + @"\FwStressFinalTestItems.xlsx";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[6, 8] = "Steps Script";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[6, 9] = s_ScriptPath_GuestNetwork + @"\FwStressSteps.xlsx";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[7, 8] = "WebGUI Script";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[7, 9] = struct_TestSettingGuestNetwork.DeivceTestScriptPath;
        //    //--------------------------------------------------------------------------------------------------//



        //    //------------------------------------- Test Result Count ------------------------------------------//
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[10, 8], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[12, 10]].Borders.LineStyle = 1;  // 畫表格框線
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[10, 8] = "Item";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[10, 9] = "PASS";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[10, 10] = "FAIL";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[11, 8] = "FW Downgrade";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[12, 8] = "FW Upgrade";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[10, 8], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[12,10]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;  // 水平置中
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[11, 8], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[12, 8]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;    // 水平靠左
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[10, 8], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[10,10]].Font.FontStyle = "Bold";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range["H10"].Font.ColorIndex = 1;  // 黑色
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range["I10"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue); ;  // 淡藍色
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range["J10"].Font.ColorIndex = 3;  // 紅色
        //    //--------------------------------------------------------------------------------------------------//



        //    //---------------------------------------- Report Data ---------------------------------------------//
        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 2] = "Item Index";
        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 3] = "Test Item";
        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 4] = "Item Section";
        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 5] = "Test Result";
        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 6] = "Comment";
        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 7] = "Log File";

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 2] = "Test Times";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 3] = "Item Index";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 4] = "Test Item";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 5] = "Item Section";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 6] = "Test Result";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 7] = "Comment";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 8] = "Log File";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 9] = "Screen Shot";

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[14, 2], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[14, 10]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 2], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 10]].Font.FontStyle = "Bold";
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 2], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[15, 10]];
        //    SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "C");

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[4, 3], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[5, 3]];
        //    SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "L");

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[6, 3], st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[11, 3]];
        //    SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "C");
        //    //--------------------------------------------------------------------------------------------------//


        //    /*** Set cells width ***/
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 1], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 1]].ColumnWidth = 2;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 2], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 2]].ColumnWidth = 15;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 3], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 3]].ColumnWidth = 15;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 4], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 4]].ColumnWidth = 20;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 5], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 5]].ColumnWidth = 15;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 6], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 6]].ColumnWidth = 15;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 7], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 7]].ColumnWidth = 15;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 8], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 8]].ColumnWidth = 15;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 9], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 9]].ColumnWidth = 15;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 10], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 10]].ColumnWidth = 15;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Range[st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[2, 11], st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[10, 11]].ColumnWidth = 15;


        //    /*** Set cells width ***/
        //    //SetExcelCellsWidth();

        //}

        //private void WriteStepTestResultToArray()
        //{
        //    //** st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest]   <-- Current Step

        //    for (STEP_SCRIPT_ST_IDX_RouterTest = 0; STEP_SCRIPT_ST_IDX_RouterTest < st_ReadDeviceWebGuiScriptData_RouterTest.Length; STEP_SCRIPT_ST_IDX_RouterTest++)   //------------ Deivce Test Script Data Array -----------//
        //    {
        //        string strCurrentStepAction = st_ReadStepsScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Action;

        //        if (strCurrentStepAction == st_ReadDeviceWebGuiScriptData_RouterTest[STEP_SCRIPT_ST_IDX_RouterTest].Procedure)
        //        {
        //        }
        //    }
        //}

        //private void WriteFinalTestResultToReport()
        //{
        //    int iStepsStartIndex = Convert.ToInt32(st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].StartIndex);
        //    int iStepsStopIndex = Convert.ToInt32(st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].StopIndex);
            
        //    for (int STEP_STRUCT_INDEX = 0; STEP_STRUCT_INDEX < st_ReadStepsScriptData_RouterTest.Length; STEP_STRUCT_INDEX++)                                     //------------------ Step Script Array------------------//
        //    {
        //        int icurStepsIndex = Convert.ToInt32(st_ReadStepsScriptData_RouterTest[STEP_STRUCT_INDEX].Index);

        //        if (icurStepsIndex >= iStepsStartIndex && icurStepsIndex <= iStepsStopIndex)  // iStepsStartIndex <=  icurStepsIndex <= iStepsStopIndex
        //        {
        //            string CurrentStepTestResult = st_ReadStepsScriptData_RouterTest[STEP_STRUCT_INDEX].TestResult;
                    
        //            if (CurrentStepTestResult == null || CurrentStepTestResult == string.Empty || CurrentStepTestResult == "")
        //            {
        //                st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].TestResult = "PASS";
        //            }
        //            else if (CurrentStepTestResult.CompareTo("FAIL") == 0)
        //            {
        //                st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].TestResult = "FAIL";
        //                //st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].Comment = st_ReadStepsScriptData_RouterTest[STEP_STRUCT_INDEX].Note;
        //                break;
        //            }
        //        }
        //    }

        //    WriteFinalTestResultExcel(st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest]);
        //}

        //private void WriteFinalTestResultExcel(FinalTestItemsScriptDataWebGuiFwUpDnGrade FinalScriptPara)
        //{
        //    string TEST_ITEM = st_ReadTestItemsScript_RouterTest[TEST_ITEMS_SCRIPT_ST_IDX_RouterTest].TestItem;
        //    string StatusCount = string.Empty;
        //    int gridROW;
        //    int gridCELL;
        //    int excelROW;
        //    int excelCELL;

        //    //string strCurTestTimes = string.Format("{0}/{1}", i_TestRun_RouterTest.ToString(), struct_TestSettingGuestNetwork.TestRun);
        //    string strCurTestTimes = string.Format("{0}", i_TestRun_RouterTest.ToString());
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[4, 9] = strCurTestTimes;  //----- Update Current Test Times

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 2] = i_TestRun_RouterTest.ToString();
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 2];
        //    SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "C");

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 3] = FinalScriptPara.ItemIndex;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 3];
        //    SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "C");

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 4] = FinalScriptPara.TestItem;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 4];
        //    SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "L");

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 5] = FinalScriptPara.ItemSection;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 5];
        //    SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "L");

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 6] = FinalScriptPara.TestResult;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 6];
        //    SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "C");
        //    if (FinalScriptPara.TestResult == "PASS")
        //    {
        //        st_ExcelObjectFinalRepor_GuestNetwork.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

        //        gridROW = TestStatusDict[TEST_ITEM].GridPassRow;
        //        gridCELL = TestStatusDict[TEST_ITEM].GridPassCell;
        //        excelROW = TestStatusDict[TEST_ITEM].ExcelPassRow;
        //        excelCELL = TestStatusDict[TEST_ITEM].ExcelPassCell;
        //        if (TEST_ITEM.CompareTo("FW Downgrade") == 0)
        //        {
        //            FwDowngrade.PassCount = FwDowngrade.PassCount + 1;
        //            StatusCount = FwDowngrade.PassCount.ToString();
        //        }
        //        else if (TEST_ITEM.CompareTo("FW Upgrade") == 0)
        //        {
        //            FwUpgrade.PassCount = FwUpgrade.PassCount + 1;
        //            StatusCount = FwUpgrade.PassCount.ToString();
        //        }
        //        st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[excelROW, excelCELL] = StatusCount;                                                   //----- Update Test Result to Excel Report
        //        Invoke(new AddDataGridViewDelegate_WebGuiFwUpDnGrade(ModifyCurrentTeatStatusToDataGridView), new object[] { gridROW, gridCELL, StatusCount });  //----- Update Test Result to GUI Data Grid View

        //    }
        //    else if (FinalScriptPara.TestResult == "FAIL")
        //    {
        //        st_ExcelObjectFinalRepor_GuestNetwork.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

        //        gridROW = TestStatusDict[TEST_ITEM].GridFailRow;
        //        gridCELL = TestStatusDict[TEST_ITEM].GridFailCell;
        //        excelROW = TestStatusDict[TEST_ITEM].ExcelFailRow;
        //        excelCELL = TestStatusDict[TEST_ITEM].ExcelFailCell;
        //        //----- Update Test Result to GUI Data Grid View -----//
        //        if (TEST_ITEM.CompareTo("FW Downgrade") == 0)
        //        {
        //            FwDowngrade.FailCount = FwDowngrade.FailCount + 1;
        //            StatusCount = FwDowngrade.FailCount.ToString();
        //        }
        //        else if (TEST_ITEM.CompareTo("FW Upgrade") == 0)
        //        {
        //            FwUpgrade.FailCount = FwUpgrade.FailCount + 1;
        //            StatusCount = FwUpgrade.FailCount.ToString();
        //        }
        //        st_ExcelObjectFinalRepor_GuestNetwork.excelApp.Cells[excelROW, excelCELL] = StatusCount;                                                   //----- Update Test Result to Excel Report
        //        Invoke(new AddDataGridViewDelegate_WebGuiFwUpDnGrade(ModifyCurrentTeatStatusToDataGridView), new object[] { gridROW, gridCELL, StatusCount });  //----- Update Test Result to GUI Data Grid View


        //        //--- When Test Fail, Add Screen Shot File Link
        //        //st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 9] = FinalScriptPara.ScreenShotPath;
        //        st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 9];
        //        st_ExcelObjectFinalRepor_GuestNetwork.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
        //        SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "C");
        //        st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Hyperlinks.Add(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, FinalScriptPara.ScreenShotPath, Type.Missing, FinalScriptPara.ScreenShotPath, "ScreenShot");
        //    }
        //    else if (FinalScriptPara.TestResult == "ERROR")
        //    {
        //        st_ExcelObjectFinalRepor_GuestNetwork.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
        //    }

        //    st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 7] = FinalScriptPara.Comment;
        //    st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 7];
        //    SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "C");


        //    ////Log File Link
        //    ////st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 8] = FinalScriptPara.Log;
        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelRange = st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Cells[i_FinalReportStartRow_GuestNetwork, 8];
        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
        //    //SettingExcelAlignment(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "H", "C");
        //    //st_ExcelObjectFinalRepor_GuestNetwork.excelWorkSheet.Hyperlinks.Add(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, FinalScriptPara.ScreenShotPath, Type.Missing, FinalScriptPara.Log, "ConsoleLog");


        //    //SettingExcelFontUSBStorage(st_ExcelObjectFinalRepor_GuestNetwork.excelRange, "Times New Roman", 11, "B", 3);
        //    SaveExcelFile_Common(st_ExcelObjectFinalRepor_GuestNetwork);  // Save Excel File (每項測完都存檔一次，避免當機檔案遺失)
        //    i_FinalReportStartRow_GuestNetwork++;
        //}

        #endregion


        //=================================================================//
        //------------------------ Invoke Function ------------------------//
        //=================================================================//
        #region Invoke Function

        //private void ClearInfoTextBoxWebGuiFwUpDnGrade()
        //{
        //    TEXTBOX_INFO.Clear();
        //}
        
        //private void ToggleWebGuiFwUpDnGradeTestTimes()
        //{
        //    labelWebGuiFwUpDnGradeTestTimes.Text = i_TestRun_RouterTest.ToString();
        //}

        //private void ModifyCurrentTeatStatusToDataGridView(int iROW, int iCELL, string dataValue)
        //{
        //    dgvWebGuiFwUpDnGradeFunctionTestDataView.Rows[iROW].Cells[iCELL].Value = dataValue;
        //}


        //-----------------  System Cursor ------------------//
        #region System Cursor
        private void ToggleGuestNetworkFunctionTestSystemCursorWait()
        {
            ToggleGuestNetworkFunctionTestSystemCursorStatus(true);
        }

        private void ToggleGuestNetworkFunctionTestSystemCursorDefault()
        {
            ToggleGuestNetworkFunctionTestSystemCursorStatus(false);
        }

        private void ToggleGuestNetworkFunctionTestSystemCursorStatus(bool Toggle)
        {
            if (Toggle == true)
            {
                btnGuestNetworkFunctionTestRun.Enabled = false;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            }
            else
            {
                btnGuestNetworkFunctionTestRun.Enabled = true;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
        }
        
        #endregion
        //---------------------------------------------------//



        //------------------ Test Controller ----------------//
        #region Test Controller
        private void ToggleGuestNetworkFunctionTestSettingTrue()
        {
            ToggleGuestNetworkFunctionTestController(true);
            //Debug.WriteLine("Toggle");
        }

        private void ToggleGuestNetworkFunctionTestController(bool Toggle)
        {
            btnGuestNetworkFunctionTestRun.Text = Toggle ? "Run" : "Stop";


            //-------- MenuItem -------//
            fileToolStripMenuItem.Enabled = Toggle;
            itemToolStripMenuItem.Enabled = Toggle;
            setupToolStripMenuItem.Enabled = Toggle;
            helpToolStripMenuItem.Enabled = Toggle;

            //----- Function Test -----//
            gbox_ModelGuestNetwork.Enabled = Toggle;
            gbox_DeviceSettingGuestNetwork.Enabled = Toggle;
            gbox_GAsettingGuestNetwork.Enabled = Toggle;
            gbox_TestSettingGuestNetwork.Enabled = Toggle;


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


            if (bWebGuiCtrlFTThreadRunning == false)
            {
                btnGuestNetworkFunctionTestRun.Enabled = false;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

                try
                {
                    if (threadWebGuiCtrlFTstopEvent.IsAlive)
                    {
                        threadWebGuiCtrlFTstopEvent.Abort();
                        threadWebGuiCtrlFTstopEvent.Join();
                    }
                    if (threadWebGuiCtrlFTscriptItem.IsAlive)
                    {
                        threadWebGuiCtrlFTscriptItem.Abort();
                        threadWebGuiCtrlFTscriptItem.Join();
                    }
                    if (threadWebGuiCtrlFT.IsAlive)
                    {
                        threadWebGuiCtrlFT.Abort();
                        threadWebGuiCtrlFT.Join();
                    }
                }
                catch { }

                TestFinishedAction_GuestNetwork();


                MessageBoxTopMost("ATE Information", "Test complete !!!");  // 將 MessageBox於桌面置頂
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                btnGuestNetworkFunctionTestRun.Enabled = true;
                bWebGuiCtrlTestComplete = true;
            }
        }
        
        #endregion
        //---------------------------------------------------//
        
        #endregion
        
#endregion

    }


}
