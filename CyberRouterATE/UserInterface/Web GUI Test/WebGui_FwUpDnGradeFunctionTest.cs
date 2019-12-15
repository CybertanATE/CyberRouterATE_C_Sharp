//---------------------------------------------------------------------------------------//
//  This code was created by CyberTan Sally Lee.                                         // 
//  File           : WebGui_FwUpDnGradeFunctionTest.cs                                   // 
//  Update         : 2018-12-07                                                          //
//  Version        : 1.0.181207                                                          //
//  Description    :                                                                     //  
//  Modified       : 2018-12-07 Initial version                                          //
//  History        : 2018-12-07 Initial version                                          //  
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
        public delegate void AddDataGridViewDelegate_WebGuiFwUpDnGrade(int iROW, int iCELL, string dataValue);
        private delegate void WebGuiFwUpDnGradeSetTextCallBack(string text, TextBox textbox);
        public delegate void showWebGuiFwUpDnGradeDelegate();

        Thread threadWebGuiFwUpDnGradeFT;
        Thread threadWebGuiFwUpDnGradeFTstopEvent;
        Thread threadWebGuiFwUpDnGradeFTscriptItem;
        bool bWebGuiFwUpDnGradeFTThreadRunning = false;
        bool bWebGuiFwUpDnGradeTestComplete = true;
        bool bWebGuiFwUpDnGradeSingleScriptItemRunning = false;

        int i_TestRunWebGuiFwUpDnGrade = 1;
        int TEST_ITEMS_STRUCT_INDEX = 0;
        int STEP_SCRIPT_STRUCT_INDEX = 0;
        int DEVICE_TEST_SCRIPT_STRUCT_INDEX = 0;
        string s_CurrentURL_WebGuiFwUpDnGrade = string.Empty;
        string s_OriginalFWver = string.Empty;
        string s_UpdatedFWver = string.Empty;

        int i_DATA_FinalROW_WebGuiFwUpDnGrade = 16;
        string s_ScriptPath_WebGuiFwUpDnGrade = System.Windows.Forms.Application.StartupPath + @"\testCondition\FwStress";
        string s_ReportPathWebGuiFwUpDnGrade = System.Windows.Forms.Application.StartupPath + @"\report";
        string s_ReportFolderFullPath_WebGuiFwUpDnGrade;
        string s_FolderNameWebGuiFwUpDnGrade = "WebGuiFwUpDnGrade_DATE";
        string s_FinalReportFileName_WebGuiFwUpDnGrade = string.Empty;
        string s_FinalReportFilePath_WebGuiFwUpDnGrade;
        string s_WebGuiReportFileName_WebGuiFwUpDnGrade = string.Empty;
        string s_WebGuiReportFilePath_WebGuiFwUpDnGrade;

        FinalTestItemsScriptDataWebGuiFwUpDnGrade[] st_ReadTestItemsScriptWebGuiFwUpDnGrade;
        StepsScriptDataWebGuiFwUpDnGrade[] st_ReadStepsScriptDataWebGuiFwUpDnGrade;
        DeviceTestScriptDataWebGuiFwUpDnGrade[] st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade;

        ModelStructWebGuiFwUpDnGrade struct_ModelStructWebGuiFwUpDnGrade = new ModelStructWebGuiFwUpDnGrade();
        DeviceSettingWebGuiFwUpDnGrade struct_DeviceSettingWebGuiFwUpDnGrade  = new DeviceSettingWebGuiFwUpDnGrade();
        TestSettingWebGuiFwUpDnGrade struct_TestSettingWebGuiFwUpDnGrade = new TestSettingWebGuiFwUpDnGrade();

        ExcelObject st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade = new ExcelObject();
        //ExcelObject st_ExcelObjectWebGuiRepor_WebGuiFwUpDnGrade = new ExcelObject();

        Dictionary<string, CurrentTestStatus> TestStatusDict = new Dictionary<string, CurrentTestStatus>();
        CurrentTestStatus FwDowngrade = new CurrentTestStatus();
        CurrentTestStatus FwUpgrade = new CurrentTestStatus();




        //**********************************************************************************//
        //------------------ Web GUI FW Up/Downgrade Function Test Event -------------------//
        //**********************************************************************************//
#region Web GUI FW Up/Downgrade Test Event

        private void fWUpgradeDowngradeTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //----- Hide All TabPage -----//
            ToggleToolStripMenuItem(false);
            this.SallyTestPage.Parent = null;        // hide SallyTestPage TabPage
            tabControl_RouterStartPage.Hide();
            
            //--- Show TabPage(needed) ---//
            tabControl_GUItest.Show();
            this.tpWebGuiFwUpDnGradeFunctionTest.Parent = this.tabControl_GUItest;            //show Function Test TabPage
  
            webGUITestToolStripMenuItem.Checked = true;
            fWUpgradeDowngradeTestToolStripMenuItem.Checked = true;
            
            //ElementVisibleSetting();
            tsslMessage.Text = tabControl_GUItest.TabPages[tabControl_GUItest.SelectedIndex].Text + " Control Panel";


            txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Text = System.Windows.Forms.Application.StartupPath + @"\testCondition\DutWebScript\EA7300_WebGuiScript_.xlsx";
            txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Select(txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Text.Length, 0);  // 檔名很長，讓游標位置靠最右邊
            txtWebGuiFwUpDnGradeFunctionTestDownfradeFwFilePath.Select(txtWebGuiFwUpDnGradeFunctionTestDownfradeFwFilePath.Text.Length, 0);            // 檔名很長，讓游標位置靠最右邊
            txtWebGuiFwUpDnGradeFunctionTestUpdrageFwFilePath.Select(txtWebGuiFwUpDnGradeFunctionTestUpdrageFwFilePath.Text.Length, 0);                // 檔名很長，讓游標位置靠最右邊
            InitWebGuiFwUpDnGradeTestDataView();

        }

        private void btnWebGuiFwUpDnGradeFunctionTestDowngradeFwFileBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogWebGUI = new OpenFileDialog();
            openFileDialogWebGUI.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            // Set filter for file extension and default file extension
            //openFileDialogWebGUI.Filter = "Excel file|*.xlsx";

            // If the file name is not an empty string open it for opening.
            if (openFileDialogWebGUI.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialogWebGUI.FileName != "")
            {
                txtWebGuiFwUpDnGradeFunctionTestDownfradeFwFilePath.Clear();
                txtWebGuiFwUpDnGradeFunctionTestDownfradeFwFilePath.Text = openFileDialogWebGUI.FileName;
                txtWebGuiFwUpDnGradeFunctionTestDownfradeFwFilePath.Select(txtWebGuiFwUpDnGradeFunctionTestDownfradeFwFilePath.Text.Length, 0); // 檔名很長，讓游標位置靠最右邊
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            }
            else if (txtWebGuiFwUpDnGradeFunctionTestDownfradeFwFilePath.Text == "" || txtWebGuiFwUpDnGradeFunctionTestDownfradeFwFilePath.Text == null)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                MessageBox.Show("Please Select a Firmware File for Downgrade Test!!");
                return;
            }
        }

        private void btnWebGuiFwUpDnGradeFunctionTestUpgradeFwFileBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogWebGUI = new OpenFileDialog();
            openFileDialogWebGUI.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            // Set filter for file extension and default file extension
            //openFileDialogWebGUI.Filter = "Excel file|*.xlsx";

            // If the file name is not an empty string open it for opening.
            if (openFileDialogWebGUI.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialogWebGUI.FileName != "")
            {
                txtWebGuiFwUpDnGradeFunctionTestUpdrageFwFilePath.Clear();
                txtWebGuiFwUpDnGradeFunctionTestUpdrageFwFilePath.Text = openFileDialogWebGUI.FileName;
                txtWebGuiFwUpDnGradeFunctionTestUpdrageFwFilePath.Select(txtWebGuiFwUpDnGradeFunctionTestUpdrageFwFilePath.Text.Length, 0); // 檔名很長，讓游標位置靠最右邊
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            }
            else if (txtWebGuiFwUpDnGradeFunctionTestUpdrageFwFilePath.Text == "" || txtWebGuiFwUpDnGradeFunctionTestUpdrageFwFilePath.Text == null)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                MessageBox.Show("Please Select a Firmware File for Upgrade Test!!");
                return;
            }
        }

        private void btnWebGuiFwUpDnGradeFunctionTestDeviceTestScriptBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogWebGUI = new OpenFileDialog();
            openFileDialogWebGUI.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\testCondition\DutWebScript"; ;
            
            // Set filter for file extension and default file extension
            openFileDialogWebGUI.Filter = "Excel file|*.xlsx";

            // If the file name is not an empty string open it for opening.
            if (openFileDialogWebGUI.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialogWebGUI.FileName != "")
            {
                txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Clear();
                txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Text = openFileDialogWebGUI.FileName;
                txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Select(txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Text.Length, 0); // 檔名很長，讓游標位置靠最右邊
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            }
            else if (txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Text == "" || txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Text == null)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                MessageBox.Show("Please Select a Script File!!");
                return;
            }
        }

        private void btnWebGuiFwUpDnGradeFunctionTestSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDebugFileDialog = new SaveFileDialog();

            saveDebugFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            saveDebugFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveDebugFileDialog.FilterIndex = 1;
            saveDebugFileDialog.RestoreDirectory = true;
            saveDebugFileDialog.FileName = "WebGuiFwUpDnGradeLog";

            if (saveDebugFileDialog.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText(saveDebugFileDialog.FileName, txtWebGuiFwUpDnGradeFunctionTestInformation.Text);
            }
        }
        
        private void btnWebGuiFwUpDnGradeFunctionTestRun_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Run Btn Checked!!");

            btnWebGuiFwUpDnGradeFunctionTestRun.Enabled = false;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            if (bWebGuiFwUpDnGradeFTThreadRunning == false && bWebGuiFwUpDnGradeTestComplete == true)
            {
                #region

                //-------------- Save Test Parameter --------------//
                SaveModelParameterWebGuiFwUpDnGradeFunctionTest();
                SaveDeviceSettingParameterWebGuiFwUpDnGradeFunctionTest();
                SaveTestSettingParameterWebGuiFwUpDnGradeFunctionTest();



                //-------------------------------------------------//
                //----------------- Start Test --------------------//
                //-------------------------------------------------//

                bWebGuiFwUpDnGradeFTThreadRunning = true;
                bWebGuiFwUpDnGradeTestComplete = false;
                ToggleWebGuiFwUpDnGradeFunctionTestController(false);  /* Disable all controller */
                txtWebGuiFwUpDnGradeFunctionTestInformation.Clear();

                //--------------- Start Test Thread ---------------//
                SetText("----------------------------- Start Device GUI ATE Test -----------------------------", txtWebGuiFwUpDnGradeFunctionTestInformation);     
                threadWebGuiFwUpDnGradeFT = new Thread(new ThreadStart(DoWebGuiFwUpDnGradeFunctionTest));
                threadWebGuiFwUpDnGradeFT.Name = "";
                threadWebGuiFwUpDnGradeFT.Start();

                //*** Use another thread to catch the stop event of test thread ***//
                threadWebGuiFwUpDnGradeFTstopEvent = new Thread(new ThreadStart(threadWebGuiFwUpDnGradeFT_StopEvent));
                threadWebGuiFwUpDnGradeFTstopEvent.Name = "";
                threadWebGuiFwUpDnGradeFTstopEvent.Start();
                #endregion
            }
            else if (bWebGuiFwUpDnGradeFTThreadRunning == true && bWebGuiFwUpDnGradeTestComplete == false)
            {
                tsslMessage.Text = "Function Test Control Panel";
                bWebGuiFwUpDnGradeFTThreadRunning = false;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "The Test will stop later. Please wait!", "ATE Information", MessageBoxButtons.OK);
            }

            /* Button release */
            //Thread.Sleep(1000);
            //System.Windows.Forms.Cursor.Current = Cursors.Default;
            //btnWebGuiFwUpDnGradeFunctionTestRun.Enabled = true;
        }

#endregion



        //**********************************************************************************//
        //------------------ Web GUI FW Up/Downgrade Function Test Module ------------------//
        //**********************************************************************************//
#region Web GUI FW Up/Downgrade Test Module

        //=================================================================//
        //------------------------- Initial Function ----------------------//
        //=================================================================//
        #region Initial Function
        
        private void InitWebGuiFwUpDnGradeTestDataView()
        {
            //----- Initial Title -----//
            dgvWebGuiFwUpDnGradeFunctionTestDataView.ColumnCount = 3;
            dgvWebGuiFwUpDnGradeFunctionTestDataView.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvWebGuiFwUpDnGradeFunctionTestDataView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvWebGuiFwUpDnGradeFunctionTestDataView.Name = "Test Data View Setting";
            dgvWebGuiFwUpDnGradeFunctionTestDataView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[0].Name = "Test Case";
            dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[1].Name = "Pass";
            dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[2].Name = "Fail";

            dgvWebGuiFwUpDnGradeFunctionTestDataView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvWebGuiFwUpDnGradeFunctionTestDataView.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[0].Width = 120;
            dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[1].Width = 70;
            dgvWebGuiFwUpDnGradeFunctionTestDataView.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvWebGuiFwUpDnGradeFunctionTestDataView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;



            //----- Initial Test Status -----//

            /** Remove the all data **/
            if (dgvWebGuiFwUpDnGradeFunctionTestDataView.RowCount > 1)
            {
                DataTable dt = (DataTable)dgvWebGuiFwUpDnGradeFunctionTestDataView.DataSource;
                dgvWebGuiFwUpDnGradeFunctionTestDataView.Rows.Clear();
                dgvWebGuiFwUpDnGradeFunctionTestDataView.DataSource = dt;
            }

            FwDowngrade.PassCount = 0;
            FwDowngrade.FailCount = 0;
            FwDowngrade.GridPassRow = 0;    // GridPass(Row,Cell) = (0,1)
            FwDowngrade.GridPassCell = 1;
            FwDowngrade.GridFailRow = 0;    // GridFail(Row,Cell) = (0,2)
            FwDowngrade.GridFailCell = 2;
            FwDowngrade.ExcelPassRow = 11;  // ExcelPass(Row,Cell) = (11,9)
            FwDowngrade.ExcelPassCell = 9;
            FwDowngrade.ExcelFailRow = 11;  // ExcelFail(Row,Cell) = (11,10)
            FwDowngrade.ExcelFailCell = 10;

            FwUpgrade.PassCount = 0;
            FwUpgrade.FailCount = 0;
            FwUpgrade.GridPassRow = 1;      // GridPass(Row,Cell) = (1,1)
            FwUpgrade.GridPassCell = 1;
            FwUpgrade.GridFailRow = 1;      // GridFail(Row,Cell) = (1,2)
            FwUpgrade.GridFailCell = 2;
            FwUpgrade.ExcelPassRow = 12;    // ExcelPass(Row,Cell) = (12,9)
            FwUpgrade.ExcelPassCell = 9;
            FwUpgrade.ExcelFailRow = 12;    // ExcelFail(Row,Cell) = (12,10)
            FwUpgrade.ExcelFailCell = 10;

            TestStatusDict["FW Downgrade"] = FwDowngrade;
            TestStatusDict["FW Upgrade"] = FwUpgrade;

            AddTeatDataToDataGridView("FW Downgrade");
            AddTeatDataToDataGridView("FW Upgrade");

        }

        private void AddTeatDataToDataGridView(string TestCase)
        {
            string PassCount = TestStatusDict[TestCase].PassCount.ToString();
            string FailCount = TestStatusDict[TestCase].FailCount.ToString();

            string[] arrayData = new string[] 
            {
                TestCase,
                PassCount,
                FailCount
            };
            dgvWebGuiFwUpDnGradeFunctionTestDataView.Rows.Add(arrayData);
        }

        private void InitialTestStatus()
        {
            //----- Initial GUI Data Grid -----//
            FwDowngrade.PassCount = 0;
            FwDowngrade.FailCount = 0;
            FwUpgrade.PassCount = 0;
            FwUpgrade.FailCount = 0;

            for (int i = 0; i < dgvWebGuiFwUpDnGradeFunctionTestDataView.RowCount - 1; i++)
            {
                for (int j = 1; j < dgvWebGuiFwUpDnGradeFunctionTestDataView.ColumnCount; j++)
                {
                    Invoke(new AddDataGridViewDelegate_WebGuiFwUpDnGrade(ModifyCurrentTeatStatusToDataGridView), new object[] { i, j, "0" });
                }
            }

            //----- Initial Excel -----//
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[FwDowngrade.ExcelPassRow, FwDowngrade.ExcelPassCell] = "0";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[FwDowngrade.ExcelFailRow, FwDowngrade.ExcelFailCell] = "0";

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[FwUpgrade.ExcelPassRow, FwUpgrade.ExcelPassCell] = "0";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[FwUpgrade.ExcelFailRow, FwUpgrade.ExcelFailCell] = "0";
        }
        
        private void TestInitialWebGuiFwUpDnGrade()
        {
            RouterTestTimer = null;
            RouterTestTimer = new Stopwatch();
            st_ReadTestItemsScriptWebGuiFwUpDnGrade = null;
            st_ReadStepsScriptDataWebGuiFwUpDnGrade = null;
            st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade = null;
            
            i_DATA_FinalROW_WebGuiFwUpDnGrade = 16;
            s_ScriptPath_WebGuiFwUpDnGrade = System.Windows.Forms.Application.StartupPath + @"\testCondition\FwStress";
            s_ReportPathWebGuiFwUpDnGrade = System.Windows.Forms.Application.StartupPath + @"\report";
            s_FolderNameWebGuiFwUpDnGrade = "WebGuiFwUpDnGrade_DATE";
            s_FinalReportFileName_WebGuiFwUpDnGrade = string.Empty;
            s_WebGuiReportFileName_WebGuiFwUpDnGrade = string.Empty;
            s_OriginalFWver = string.Empty;
            s_UpdatedFWver = string.Empty;

            InitialTestStatus();

            csbt_TestBrowserType = csbt_TestBrowserChrome;
            if (!WebDriverInitial())
            {
                CloseExcelReportFile(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade);
                ClodeWebDriver();
                bCommonFTThreadRunning = false;
                bWebGuiFwUpDnGradeSingleScriptItemRunning = false;
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
        
        private void threadWebGuiFwUpDnGradeFT_StopEvent()
        {
            while (bWebGuiFwUpDnGradeFTThreadRunning == true)
            {
                Thread.Sleep(500);
            }

            WebGuiFwUpDnGradeFunctionTestThreadFinished();
        }

        private void WebGuiFwUpDnGradeFunctionTestThreadFinished()
        {
            Invoke(new WebGuiFwUpDnGradeSetTextCallBack(SetText), new object[] { "------------------------------ Device GUI ATE Test End ------------------------------", txtWebGuiFwUpDnGradeFunctionTestInformation }); 
            this.Invoke(new showWebGuiFwUpDnGradeDelegate(ToggleWebGuiFwUpDnGradeFunctionTestSettingTrue));
        }

        private void TestFinishedActionWebGuiFwUpDnGrade()
        {
            CloseExcelReportFile(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade);
            ClodeWebDriver();
            Thread.Sleep(3000);
        }

        private void CloseExcelReportFile(ExcelObject excelObject)
        {
            try
            {
                SaveEndTimetoExcel_Common(excelObject, 1, 7, 3);
                SetExcelCellsHigh_Common(excelObject, "B16", "B50000");
                Save_and_Close_Excel_File_Common(excelObject);
            }
            catch { }
        }

        //private void ClodeWebDriver()
        //{
        //    try
        //    {
        //        bool bResult = cs_DutWebGuiCtrlClass.Close_WebDriver();
        //        cs_DutWebGuiCtrlClass = null;
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
        //        checkXPathResult = cs_DutWebGuiCtrlClass.CheckXPathDisplayed(ScriptPara.ElementXpath);
        //        Thread.Sleep(100);

        //    } while (waitingtime.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);
        //    waitingtime.Stop();
        //    Thread.Sleep(2000);

        //    return checkXPathResult;
        //}

        private bool InTimePeriod()
        {
            string sInfo = string.Empty;
            DateTime SettingStartDay = dtpWebGuiFwUpDnGradeDeviceSettingPeriodStartDay.Value;
            DateTime SettingEndDay = dtpWebGuiFwUpDnGradeDeviceSettingPeriodEndDay.Value;

            //--- 省掉日期資訊，只關心時間 ---//
            string SystemTimeNow               = DateTime.Now.ToString("HH:mm:ss");
            string SystemTimeNowContainRunTime = DateTime.Now.AddMinutes(20).ToString("HH:mm:ss");  // 目前時間再加上粗估執行一次的時間(20分鐘)
            string SettingStartTime = SettingStartDay.ToString("HH:mm:ss");
            string SettingEndTime   = SettingEndDay.ToString("HH:mm:ss");

            DateTime dtStartTime = Convert.ToDateTime(SettingStartTime);
            DateTime dtEndTime   = Convert.ToDateTime(SettingEndTime);
            DateTime dtTimeNow               = Convert.ToDateTime(SystemTimeNow);
            DateTime dtTimeNowContainRunTime = Convert.ToDateTime(SystemTimeNowContainRunTime);

            //int i = DateTime.Compare(dt1, dt3); // dt1 比 dt3大 --> 1
            //int k = DateTime.Compare(dt2, dt3); // dt2 比 dt3小 --> -1
            if ((DateTime.Compare(dtTimeNow, dtStartTime) >= 0 && DateTime.Compare(dtTimeNow, dtEndTime) <= 0) ||
                 (DateTime.Compare(dtTimeNowContainRunTime, dtStartTime) >= 0 && DateTime.Compare(dtTimeNowContainRunTime, dtEndTime) <= 0))
            {
                return true;
            }

            return false;
        }
        
        private void SaveWebGuiTestResult(ref CBT_SeleniumApi.GuiScriptParameter ScriptPara, bool bTestResult)
        {
            if (ScriptPara.TestResult == null || ScriptPara.TestResult == "")
            {
                if (bTestResult == true || ScriptPara.Note == string.Empty || ScriptPara.Note == "")
                    ScriptPara.TestResult = "PASS";
                else
                    ScriptPara.TestResult = "FAIL";
            }

            //***** WriteStepTestResultToArray *****//
            //-- st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX]   <-- Current Step
            string CuurentStepTestResult = st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].TestResult;
            if (ScriptPara.TestResult.CompareTo("FAIL") == 0)
            {
                st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].TestResult = "FAIL";
                st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Note = ScriptPara.Note;
            }
        }
        
        private void DoWebGuiFwUpDnGradeFunctionTest()
        {
           // MessageBox.Show("DoWebGuiFwUpDnGradeFunctionTest()");

            string sInfo = string.Empty;
            string strCurTestTimes = string.Empty;
            int TestTimeOut = 10000;
            int iTestRun = Convert.ToInt32(struct_TestSettingWebGuiFwUpDnGrade.TestRun);

            this.Invoke(new showWebGuiFwUpDnGradeDelegate(ToggleWebGuiFwUpDnGradeFunctionTestSystemCursorWait));

            CreateNewFinalReport_WebGuiFwUpDnGrade();
            //CreateNewWebGuiReport_WebGuiFwUpDnGrade();
            TestInitialWebGuiFwUpDnGrade();

            sInfo = string.Format("Loading Test Items Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { sInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
            WebGuiFwUpDnGradeFunctionTestLoadFinalTestItemsScript();

            sInfo = string.Format("Loading Steps Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { sInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
            WebGuiFwUpDnGradeFunctionTestLoadStepsScript();

            sInfo = string.Format("Loading WebGUI Script... ");
            Invoke(new SetTextCallBack(SetText), new object[] { sInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
            WebGuiFwUpDnGradeFunctionTestLoadWebGUIScript();

            
            for (i_TestRunWebGuiFwUpDnGrade = 1; i_TestRunWebGuiFwUpDnGrade <= iTestRun; i_TestRunWebGuiFwUpDnGrade++)
            {
              //DateTime Now = DateTime.Now;

                //this.Invoke(new showWebGuiFwUpDnGradeDelegate(ClearInfoTextBoxWebGuiFwUpDnGrade));

                sInfo = string.Format("****************** [ Test Times: {0} ] ******************", i_TestRunWebGuiFwUpDnGrade.ToString());
                //strCurTestTimes = string.Format("{0}/{1}", i_TestRunWebGuiFwUpDnGrade.ToString(), struct_TestSettingWebGuiFwUpDnGrade.TestRun);
                //strCurTestTimes = string.Format("{0}", i_TestRunWebGuiFwUpDnGrade.ToString());
                //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[4, 9] = strCurTestTimes;  //----- Update Current Test Times to Excel
                Invoke(new SetTextCallBack(SetText), new object[] { "", txtWebGuiFwUpDnGradeFunctionTestInformation });
                Invoke(new SetTextCallBack(SetText), new object[] { sInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });

                this.Invoke(new showWebGuiFwUpDnGradeDelegate(ToggleWebGuiFwUpDnGradeTestTimes));
                this.Invoke(new showWebGuiFwUpDnGradeDelegate(ToggleWebGuiFwUpDnGradeFunctionTestSystemCursorDefault));


                //---------------------------------------------------------------//
                //---- 判斷目前時間是否在 FW Check Time 範圍內，是的話則不做測試 ----//
                //---------------------------------------------------------------//
                DateTime SettingStartDay = dtpWebGuiFwUpDnGradeDeviceSettingPeriodStartDay.Value;
                DateTime SettingEndDay = dtpWebGuiFwUpDnGradeDeviceSettingPeriodEndDay.Value;
                string SettingStartTime = SettingStartDay.ToString("HH:mm:ss");
                string SettingEndTime = SettingEndDay.ToString("HH:mm:ss");

                if (InTimePeriod())
                {
                    sInfo = string.Format("<< Suspend Testing >> Wait for FW Check... (From {0} to {1})", SettingStartTime, SettingEndTime);
                    Invoke(new SetTextCallBack(SetText), new object[] { sInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });

                    while (InTimePeriod())
                    {
                        Thread.Sleep(5000);
                    }
                }
                
                
                for (TEST_ITEMS_STRUCT_INDEX = 0; TEST_ITEMS_STRUCT_INDEX < st_ReadTestItemsScriptWebGuiFwUpDnGrade.Length; TEST_ITEMS_STRUCT_INDEX++)                                            //------------------ Final Script Array-----------------//
                {
                    if (!WebDriverInitial())   //------------- FW Stress 每次 Up/Downgrade 完都要重新開啟 Web Driver
                    {
                        CloseExcelReportFile(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade);
                        ClodeWebDriver();
                        sInfo = string.Format("### FAIL in :  <Test Item> {0}  \n### Note: WebDriver Initial() FAIL!!", st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].TestItem);
                        st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].TestResult = "FAIL";
                        st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].Comment = sInfo;
                        Invoke(new SetTextCallBack(SetText), new object[] { "", txtWebGuiFwUpDnGradeFunctionTestInformation });
                        Invoke(new SetTextCallBack(SetText), new object[] { sInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
                        WriteFinalTestResultExcel(st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX]);
                        continue;
                    }

                    s_OriginalFWver = string.Empty;
                    s_UpdatedFWver = string.Empty;

                    int iStepsStartIndex = Convert.ToInt32(st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].StartIndex);
                    int iStepsStopIndex = Convert.ToInt32(st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].StopIndex);
                    sInfo = string.Format("***** <Test Item> : {0} *****", st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].TestItem);
                    Invoke(new SetTextCallBack(SetText), new object[] { "", txtWebGuiFwUpDnGradeFunctionTestInformation });
                    Invoke(new SetTextCallBack(SetText), new object[] { sInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });

                    bool bSkipTestStep = false;

                    for (STEP_SCRIPT_STRUCT_INDEX = 0; STEP_SCRIPT_STRUCT_INDEX < st_ReadStepsScriptDataWebGuiFwUpDnGrade.Length; STEP_SCRIPT_STRUCT_INDEX++)                                     //------------------ Step Script Array------------------//
                    {
                        int icurStepsIndex = Convert.ToInt32(st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Index);

                        if (bSkipTestStep == false && icurStepsIndex >= iStepsStartIndex && icurStepsIndex <= iStepsStopIndex)  // iStepsStartIndex <=  icurStepsIndex <= iStepsStopIndex
                        {
                            string sStepName = st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Name;

                            sInfo = string.Format("***    <Test Steps>: {0}", sStepName);
                            Invoke(new SetTextCallBack(SetText), new object[] { "", txtWebGuiFwUpDnGradeFunctionTestInformation });
                            Invoke(new SetTextCallBack(SetText), new object[] { sInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });

                            for (DEVICE_TEST_SCRIPT_STRUCT_INDEX = 0; DEVICE_TEST_SCRIPT_STRUCT_INDEX < st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade.Length; DEVICE_TEST_SCRIPT_STRUCT_INDEX++)   //------------ Deivce Test Script Data Array -----------//
                            {
                                string strCurrentStepAction = st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Action;

                                if (strCurrentStepAction == st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].Procedure)
                                {
                                    #region Device Test Script Main Test Thread

                                    if (st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].TestTimeOut == null || st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].TestTimeOut == "")
                                    {
                                        st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].TestTimeOut = struct_TestSettingWebGuiFwUpDnGrade.DefaultTimeout;
                                        TestTimeOut = Convert.ToInt32(struct_TestSettingWebGuiFwUpDnGrade.DefaultTimeout) * 1000;
                                    }
                                    else
                                    {
                                        TestTimeOut = Convert.ToInt32(st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].TestTimeOut) * 1000;
                                    }


                                    //----------------------------------------------------------------//
                                    //----------------------- Main Test Thread -----------------------//
                                    //----------------------------------------------------------------//              
                                    RouterTestTimer.Reset();
                                    RouterTestTimer.Start();
                                    bWebGuiFwUpDnGradeSingleScriptItemRunning = true;
                                    threadWebGuiFwUpDnGradeFTscriptItem = new Thread(new ThreadStart(WebGuiFwUpDnGradeFunctionTestMainFunction));
                                    //threadWebGuiFwUpDnGradeFTscriptItem.Name = "";
                                    threadWebGuiFwUpDnGradeFTscriptItem.Start();

                                    while (RouterTestTimer.ElapsedMilliseconds <= TestTimeOut && bWebGuiFwUpDnGradeSingleScriptItemRunning == true)
                                    {
                                        Thread.Sleep(100);
                                    }
                                    RouterTestTimer.Stop();

                                    //----------------------------------------------//
                                    //------------------ Time Out ------------------//
                                    //----------------------------------------------//
                                    if (RouterTestTimer.ElapsedMilliseconds > TestTimeOut)
                                    {
                                        #region Time Out

                                        st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Note = "Time Out!!   " + st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Note;

                                        bSkipTestStep = true;
                                        break;
                                        //if (threadWebGuiFwUpDnGradeFTscriptItem.IsAlive)
                                        //{
                                        //    threadWebGuiFwUpDnGradeFTscriptItem.Abort();
                                        //    threadWebGuiFwUpDnGradeFTscriptItem.Join();
                                        //}
                                        //bWebGuiFwUpDnGradeFTThreadRunning = false;

                                        #endregion
                                    }
                                    else if (st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].TestResult == "FAIL")
                                    {
                                        bSkipTestStep = true;
                                        break;
                                    }

                                    //WriteWebGuiTestResultToReport();        /////////////////////////////////// Write Web GUI Test Report
                                    #endregion
                                }
                            }

                            if (bSkipTestStep == true)
                            {
                                #region When bSkipTestStep == true

                                string failInfo = string.Format("### FAIL in: <Test Item> {0} (item index:{1}),   <Test Steps> {2} (step index:{3}),   <Action Name> {4}", st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].TestItem, st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].ItemIndex, st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Name, st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Index, st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].ActionName);
                                Invoke(new SetTextCallBack(SetText), new object[] { "", txtWebGuiFwUpDnGradeFunctionTestInformation });
                                Invoke(new SetTextCallBack(SetText), new object[] { failInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });

                                string failNote = string.Format("### Note: {0}", st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Note);
                                Invoke(new SetTextCallBack(SetText), new object[] { failNote, txtWebGuiFwUpDnGradeFunctionTestInformation });

                                //------ Save FAIL Comment ------//
                                st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].Comment = failInfo + "\n" + failNote;


                                //------ ScreenShot ------//
                                string TestItem = st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].TestItem;
                                string RunTimes = i_TestRunWebGuiFwUpDnGrade.ToString();
                                string DateTimeStr = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                              
                                string ScreenShotFolderPath = s_ReportFolderFullPath_WebGuiFwUpDnGrade + @"\ScreenShot";
                                bool isFolderExists = System.IO.Directory.Exists(ScreenShotFolderPath);
                                if (!isFolderExists)
                                    System.IO.Directory.CreateDirectory(ScreenShotFolderPath);

                                string ScreenShotFilePath = string.Format(ScreenShotFolderPath + @"\{0}_Run{1}_{2}.jpg", TestItem, RunTimes, DateTimeStr);
                                PrintFullScreen_Common(ScreenShotFilePath);
                                st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].ScreenShotPath = ScreenShotFilePath;

                                #endregion
                            }

                        }

                        //WriteStepTestResultToArray();        /////////////////////////////////// Write Step Test Result
                    }

                    ClodeWebDriver();  //------------- FW Stress 每次 Up/Downgrade 完都要 Close Web Driver
                    WriteFinalTestResultToReport();        /////////////////////////////////// Write Final Test Result
                }

            }

            bWebGuiFwUpDnGradeFTThreadRunning = false;
        }

        private void WebGuiFwUpDnGradeFunctionTestMainFunction()
        {
            //MessageBox.Show("WebGuiFwUpDnGradeFunctionTestMainFunction()");

            Thread.Sleep(2000);
            CBT_SeleniumApi.GuiScriptParameter ScriptPara = new CBT_SeleniumApi.GuiScriptParameter();
            string sLoginURL = string.Format("http://{0}:{1}", struct_DeviceSettingWebGuiFwUpDnGrade.DeviceIP, struct_DeviceSettingWebGuiFwUpDnGrade.HTTP_Port);
            string sExceptionInfo = string.Empty;
            string strInfo = string.Empty;
            bool bTestResult = false;

            ScriptPara.Procedure = st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].Procedure;
            ScriptPara.Index = st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].TestIndex;
            ScriptPara.Action = st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].Action;
            ScriptPara.ActionName = st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].ActionName;
            ScriptPara.ElementType = st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].ElementType;
            ScriptPara.ElementXpath = st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].ElementXpath;
            ScriptPara.ElementXpath = ScriptPara.ElementXpath.Replace('\"', '\'');
            ScriptPara.RadioBtnExpectedValueXpath = st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].RadioButtonExpectedValueXpath.Replace('\"', '\'');
            ScriptPara.WriteValue = st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].WriteExpectedValue;
            ScriptPara.ExpectedValue = st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].WriteExpectedValue;
            ScriptPara.URL = sLoginURL + st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].WriteExpectedValue;
            ScriptPara.TestTimeOut = st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].TestTimeOut;
            ScriptPara.Note = string.Empty;


            strInfo = string.Format("**      [WebGUI Index]: {0}    [Action]:{1}    [Action Name]:{2}    ", ScriptPara.Index, ScriptPara.Action, ScriptPara.ActionName);
            Invoke(new SetTextCallBack(SetText), new object[] { strInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
            strInfo = string.Empty;

            //---------------------------------------//
            //-------------- Go To URL --------------//
            //---------------------------------------//
            #region Go To URL
            //if (DEVICE_TEST_SCRIPT_STRUCT_INDEX == 0)
            if (ScriptPara.Action.CompareTo("Goto") == 0 && ScriptPara.ActionName.CompareTo("Go to Login Page") == 0)
            {
                s_CurrentURL_WebGuiFwUpDnGrade = sLoginURL;
                bTestResult = cs_DutWebGuiCtrlClass.GoToURL(sLoginURL);
                Thread.Sleep(3000);
            }
            else if (ScriptPara.Action.CompareTo("Goto") == 0 && s_CurrentURL_WebGuiFwUpDnGrade.CompareTo(ScriptPara.URL) != 0)
            {
                s_CurrentURL_WebGuiFwUpDnGrade = ScriptPara.URL;
                bTestResult = cs_DutWebGuiCtrlClass.GoToURL(s_CurrentURL_WebGuiFwUpDnGrade);
                Thread.Sleep(3000);
            }
            #endregion


            //---------------------------------------//
            //-------------- Set Value --------------//
            //---------------------------------------//
            #region Set Value
            else if (ScriptPara.Action.CompareTo("Set") == 0)
            {
                bool checkXPathResult = false;

                RouterTestTimer.Stop();
                checkXPathResult = CheckWebGuiXPath(ref ScriptPara);
                RouterTestTimer.Start();

                if (checkXPathResult == true)
                {
                    try
                    {
                        bTestResult = cs_DutWebGuiCtrlClass.SetWebElementValue(ref ScriptPara);
                    }
                    catch (Exception ex)
                    {
                        //ExceptionActionUSBStorage(s_CurrentURLUSBStorage, ref ScriptPara);
                        strInfo = "-----> Set Value Error:\n" + ex.ToString();
                        ScriptPara.Note = strInfo;
                        bTestResult = false;
                        Invoke(new SetTextCallBackT(SetText), new object[] { strInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
                        //SwitchToRGconsoleUSBStorage();
                        //WriteconsoleLogUSBStorage();
                    }
                }
                else
                {
                    strInfo = string.Format("-----> Couldn't find the Web GUI element!");
                    ScriptPara.Note = strInfo;
                    bTestResult = false;
                    Invoke(new SetTextCallBack(SetText), new object[] { strInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
                }
            }
            #endregion


            //---------------------------------------//
            //-------------- Get Value --------------//
            //---------------------------------------//
            #region Get Value
            else if (ScriptPara.Action.CompareTo("Get") == 0)
            {
                ScriptPara.Note = string.Empty;
                ScriptPara.GetValue = string.Empty;

                bool checkXPathResult = false;

                RouterTestTimer.Stop();
                checkXPathResult = CheckWebGuiXPath(ref ScriptPara);
                RouterTestTimer.Start();

                if (checkXPathResult == true)
                {
                    if (ScriptPara.ElementType.CompareTo("RADIO_BUTTON") == 0)
                    {
                        ScriptPara.ElementXpath = ScriptPara.RadioBtnExpectedValueXpath;
                    }

                    try
                    {
                        Thread.Sleep(2000);
                        bTestResult = cs_DutWebGuiCtrlClass.GetWebElementValue(ref ScriptPara);
                    }
                    catch (Exception ex)
                    {
                        strInfo = "-----> Get Value Error:\n" + ex.ToString();
                        ScriptPara.Note = strInfo;
                        bTestResult = false;
                        Invoke(new SetTextCallBackT(SetText), new object[] { strInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
                    }
                }
                else
                {
                    strInfo = string.Format("Couldn't find the Web GUI element!");
                    ScriptPara.Note = strInfo;
                    bTestResult = false;
                    Invoke(new SetTextCallBack(SetText), new object[] { strInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
                }


                //-------- Save FW version --------//
                if (ScriptPara.ActionName.CompareTo("Get original FW version") == 0)
                {
                    s_OriginalFWver = ScriptPara.GetValue;
                    strInfo = string.Format("-----> Current FW version: {0}", s_OriginalFWver);
                }
                else if (ScriptPara.ActionName.CompareTo("Get updated FW version") == 0)
                {
                    s_UpdatedFWver = ScriptPara.GetValue;

                    if (s_UpdatedFWver != null && s_UpdatedFWver != "" && s_UpdatedFWver.CompareTo(s_OriginalFWver) != 0)
                    {
                        ScriptPara.TestResult = "PASS";
                        strInfo = string.Format("-----> Updated FW version:{0}", s_UpdatedFWver);
                    }
                    else if (s_UpdatedFWver.CompareTo(s_OriginalFWver) == 0)
                    {
                        ScriptPara.TestResult = "FAIL";
                        strInfo = string.Format("-----> Original FW version:{0}\nUpdated FW version:{1}", s_OriginalFWver, s_UpdatedFWver);
                        ScriptPara.Note = strInfo;
                    }
                    else
                    {
                        ScriptPara.TestResult = "FAIL";
                        strInfo = "-----> Get updated FW version Fail.";
                        ScriptPara.Note = strInfo;
                    }
                }
                Invoke(new SetTextCallBack(SetText), new object[] { strInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
                
            }
            #endregion


            //---------------------------------------//
            //------------- File Upload -------------//
            //---------------------------------------//
            #region File Upload
            else if (ScriptPara.Action.CompareTo("FileUpload") == 0)
            {
                ScriptPara.Note = string.Empty;

                if (ScriptPara.WriteValue.CompareTo("<DowngradeFwPath>") == 0)
                    ScriptPara.WriteValue = struct_TestSettingWebGuiFwUpDnGrade.DowngradeFwFilePath;
                else if (ScriptPara.WriteValue.CompareTo("<UpgradeFwPath>") == 0)
                    ScriptPara.WriteValue = struct_TestSettingWebGuiFwUpDnGrade.UpgradeFwFilePath;

                bool checkXPathResult = false;

                RouterTestTimer.Stop();
                checkXPathResult = CheckWebGuiXPath(ref ScriptPara);
                RouterTestTimer.Start();

                if (checkXPathResult == true)
                {
                    try
                    {
                        cs_DutWebGuiCtrlClass.FileUpload(ref ScriptPara);
                    }
                    catch (Exception ex)
                    {
                        //ExceptionActionUSBStorage(cs_DutWebGuiCtrlClass, ref ScriptPara);
                        ScriptPara.Note = "-----> File Upload Error:\n" + ex.ToString();
                        //SwitchToRGconsoleUSBStorage();
                        //WriteconsoleLogUSBStorage();
                    }
                }
                else
                {
                    strInfo = string.Format("-----> Couldn't find the File Upload Butten!");
                    ScriptPara.Note = strInfo;
                    bTestResult = false;
                    Invoke(new SetTextCallBack(SetText), new object[] { strInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
                }
            }
            #endregion


            //---------------------------------------//
            //----------------- Wait ----------------//
            //---------------------------------------//
            #region Wait
            else if (ScriptPara.Action.CompareTo("Wait") == 0)
            {
                ScriptPara.Note = string.Empty;
                if (ScriptPara.ElementXpath.CompareTo("") == 0)
                {
                    strInfo = string.Format("Waiting for {0} seconds...", ScriptPara.TestTimeOut);
                    Invoke(new SetTextCallBack(SetText), new object[] { strInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });

                    RouterTestTimer.Stop();
                    Thread.Sleep(Convert.ToInt32(ScriptPara.TestTimeOut) * 1000);
                    RouterTestTimer.Start();
                    //SwitchToRGconsoleUSBStorage();
                    //WriteconsoleLogUSBStorage();
                }
                else
                {
                    bool checkXPathResult = false;

                    RouterTestTimer.Stop();
                    checkXPathResult = CheckWebGuiXPath(ref ScriptPara);
                    RouterTestTimer.Start();


                    if (checkXPathResult == false)
                    {
                        strInfo = string.Format("-----> Couldn't find the Web GUI element!");
                        ScriptPara.Note = strInfo;
                        bTestResult = false;
                        Invoke(new SetTextCallBack(SetText), new object[] { strInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });
                    }


                    //SwitchToRGconsoleUSBStorage();
                    //WriteconsoleLogUSBStorage();
                }
            }
            #endregion


            //---------------------------------------//
            //---------------- Alert ----------------//
            //---------------------------------------//
            #region Alert
            else if (ScriptPara.Action.CompareTo("Alert") == 0)
            {
            }
            #endregion


            //---------------------------------------//
            //-------------- Close Web --------------//
            //---------------------------------------//
            #region Close Web
            else if (ScriptPara.Action.CompareTo("CloseWeb") == 0)
            {
                ClodeWebDriver();
            }
            #endregion
            


            //----- Save Current Test Result -----//
            SaveWebGuiTestResult(ref ScriptPara, bTestResult);


            bWebGuiFwUpDnGradeSingleScriptItemRunning = false;
        }

        private void ExceptionAction()
        {
            string failInfo = string.Format("### FAIL in: <Test Item> {0} (item index:{1}),   <Test Steps> {2} (step index:{3}),   <Action Name> {4}", st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].TestItem, st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].ItemIndex, st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Name, st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Index, st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].ActionName);
            Invoke(new SetTextCallBack(SetText), new object[] { "", txtWebGuiFwUpDnGradeFunctionTestInformation });
            Invoke(new SetTextCallBack(SetText), new object[] { failInfo, txtWebGuiFwUpDnGradeFunctionTestInformation });

            string failNote = string.Format("### Note: {0}", st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Note);
            Invoke(new SetTextCallBack(SetText), new object[] { failNote, txtWebGuiFwUpDnGradeFunctionTestInformation });

            //------ Save FAIL Comment ------//
            st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].Comment = failInfo + "\n" + failNote;


            //------ ScreenShot ------//
            string TestItem = st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].TestItem;
            string RunTimes = i_TestRunWebGuiFwUpDnGrade.ToString();
            string DateTimeStr = DateTime.Now.ToString("yyyyMMdd-HHmmss");

            string ScreenShotFolderPath = s_ReportFolderFullPath_WebGuiFwUpDnGrade + @"\ScreenShot";
            bool isFolderExists = System.IO.Directory.Exists(ScreenShotFolderPath);
            if (!isFolderExists)
                System.IO.Directory.CreateDirectory(ScreenShotFolderPath);

            string ScreenShotFilePath = string.Format(ScreenShotFolderPath + @"\{0}_Run{1}_{2}.jpg", TestItem, RunTimes, DateTimeStr);
            PrintFullScreen_Common(ScreenShotFilePath);
            st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].ScreenShotPath = ScreenShotFilePath;
        }

        #endregion


        //=================================================================//
        //----------------------- Save Test Parameter ---------------------//
        //=================================================================//
        #region Save Test Parameter

        private void SaveModelParameterWebGuiFwUpDnGradeFunctionTest()
        {
            struct_ModelStructWebGuiFwUpDnGrade.ModelName = txtWebGuiFwUpDnGradeFunctionTestModelName.Text;
            struct_ModelStructWebGuiFwUpDnGrade.SN = txtWebGuiFwUpDnGradeFunctionTestModelSerialNumber.Text;
            struct_ModelStructWebGuiFwUpDnGrade.SWver = txtWebGuiFwUpDnGradeFunctionTestModelSwVersion.Text;
            struct_ModelStructWebGuiFwUpDnGrade.HWver = txtWebGuiFwUpDnGradeFunctionTestModelHwVersion.Text;
        }

        private void SaveDeviceSettingParameterWebGuiFwUpDnGradeFunctionTest()
        {
            struct_DeviceSettingWebGuiFwUpDnGrade.DeviceIP = masktxtWebGuiFwUpDnGradeFunctionTestDeviceIP.Text;
            struct_DeviceSettingWebGuiFwUpDnGrade.HTTP_Port = nudWebGuiFwUpDnGradeFunctionTestDefaultSettingsHTTPport.Value.ToString();

            if (ckboxWebGuiFwUpDnGradeFunctionTestConsoleLog.Checked)
                struct_DeviceSettingWebGuiFwUpDnGrade.ConsoleLog = true;
            else
                struct_DeviceSettingWebGuiFwUpDnGrade.ConsoleLog = false;
        }

        private void SaveTestSettingParameterWebGuiFwUpDnGradeFunctionTest()
        {
            struct_TestSettingWebGuiFwUpDnGrade.DowngradeFwFilePath = txtWebGuiFwUpDnGradeFunctionTestDownfradeFwFilePath.Text;
            struct_TestSettingWebGuiFwUpDnGrade.UpgradeFwFilePath = txtWebGuiFwUpDnGradeFunctionTestUpdrageFwFilePath.Text;
            struct_TestSettingWebGuiFwUpDnGrade.DeivceTestScriptPath = txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Text;
            struct_TestSettingWebGuiFwUpDnGrade.TestRun = nudWebGuiFwUpDnGradeFunctionTestTestRun.Value.ToString();
            struct_TestSettingWebGuiFwUpDnGrade.DefaultTimeout = nudWebGuiFwUpDnGradeFunctionTestDefaultTimeout.Value.ToString();

            if (ckboxWebGuiFwUpDnGradeFunctionTestStopWhenTestError.Checked)
                struct_TestSettingWebGuiFwUpDnGrade.StopWhenTestError = true;
            else
                struct_TestSettingWebGuiFwUpDnGrade.StopWhenTestError = false;
        }

        #endregion


        //=================================================================//
        //------------------------ Load Test Script -----------------------//
        //=================================================================//
        #region Load Test Script

        private void WebGuiFwUpDnGradeFunctionTestLoadFinalTestItemsScript()
        {
            string filePath = System.Windows.Forms.Application.StartupPath + @"\testCondition\FwStress\FwStressFinalTestItems.xlsx";
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
            st_ReadTestItemsScriptWebGuiFwUpDnGrade = new FinalTestItemsScriptDataWebGuiFwUpDnGrade[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadTestItemsScriptWebGuiFwUpDnGrade[SCRIPT_INDEX].ItemIndex = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadTestItemsScriptWebGuiFwUpDnGrade[SCRIPT_INDEX].TestItem = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadTestItemsScriptWebGuiFwUpDnGrade[SCRIPT_INDEX].ItemSection = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadTestItemsScriptWebGuiFwUpDnGrade[SCRIPT_INDEX].StartIndex = ScriptDataArray[SCRIPT_INDEX, 3];
                st_ReadTestItemsScriptWebGuiFwUpDnGrade[SCRIPT_INDEX].StopIndex = ScriptDataArray[SCRIPT_INDEX, 4];
                st_ReadTestItemsScriptWebGuiFwUpDnGrade[SCRIPT_INDEX].TestResult = "";
                st_ReadTestItemsScriptWebGuiFwUpDnGrade[SCRIPT_INDEX].Comment = "";
                st_ReadTestItemsScriptWebGuiFwUpDnGrade[SCRIPT_INDEX].Log = "";
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal);   // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
        }
        
        private void WebGuiFwUpDnGradeFunctionTestLoadStepsScript()
        {
            string filePath = System.Windows.Forms.Application.StartupPath + @"\testCondition\FwStress\FwStressSteps.xlsx";
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
            st_ReadStepsScriptDataWebGuiFwUpDnGrade = new StepsScriptDataWebGuiFwUpDnGrade[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadStepsScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].Index = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadStepsScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].FunctionName = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadStepsScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].Name = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadStepsScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].Command = new string[20];

                int commandlistIndex = 0;
                for (int DC = 3; DC < ColumnCount; DC++)
                {
                    if (ScriptDataArray[SCRIPT_INDEX, DC] != null && ScriptDataArray[SCRIPT_INDEX, DC] != "")
                    {
                        string tmpData = ScriptDataArray[SCRIPT_INDEX, DC].Replace("::", "$");
                        switch (tmpData.Split('$')[0].Trim())
                        {
                            case "Action":
                                st_ReadStepsScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].Action = tmpData.Split('$')[1].Trim();
                                break;

                            case "Command":
                                st_ReadStepsScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].Command[commandlistIndex] = tmpData.Split('$')[1].Trim();
                                commandlistIndex++;
                                break;

                            //case "WriteValue":
                            //    st_ReadStepsScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].WriteValue = tmpData.Split('$')[1].Trim();
                            //    break;

                            case "ExpectedValue":
                                st_ReadStepsScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].ExpectedValue = tmpData.Split('$')[1].Trim();
                                break;

                            case "TimeOut":
                                st_ReadStepsScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].TestTimeOut = tmpData.Split('$')[1].Trim();
                                break;
                        }
                    }
                }
            }
            //progressbarForm.Dispose();
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            File.SetAttributes(filePath, FileAttributes.Normal);    // 屬性設定成可讀寫，避免不正常結束造成檔案唯讀
        }

        private void WebGuiFwUpDnGradeFunctionTestLoadWebGUIScript()
        {
            string filePath = txtWebGuiFwUpDnGradeFunctionTestDeviceTestScriptFilePath.Text;
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
            st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade = new DeviceTestScriptDataWebGuiFwUpDnGrade[ActualDataCount];

            for (int SCRIPT_INDEX = 0; SCRIPT_INDEX < ActualDataCount; SCRIPT_INDEX++)
            {
                st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].Procedure = ScriptDataArray[SCRIPT_INDEX, 0];
                st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].TestIndex = ScriptDataArray[SCRIPT_INDEX, 1];
                st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].Action = ScriptDataArray[SCRIPT_INDEX, 2];
                st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].ActionName = ScriptDataArray[SCRIPT_INDEX, 3];
                st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].ElementType = ScriptDataArray[SCRIPT_INDEX, 4];
                st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].WriteExpectedValue = ScriptDataArray[SCRIPT_INDEX, 5];
                st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].RadioButtonExpectedValueXpath = ScriptDataArray[SCRIPT_INDEX, 6];
                st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].ElementXpath = ScriptDataArray[SCRIPT_INDEX, 7];
                st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].TestTimeOut = ScriptDataArray[SCRIPT_INDEX, 8];
                st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[SCRIPT_INDEX].GetValue = "";
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

        private void CreateNewFinalReport_WebGuiFwUpDnGrade()
        {
            s_FinalReportFileName_WebGuiFwUpDnGrade = "MODEL_FinalReport_SN_SW_HW_DATE.xlsx";
            s_FinalReportFileName_WebGuiFwUpDnGrade = getReportFileNameWebGuiFwUpDnGrade(s_FinalReportFileName_WebGuiFwUpDnGrade);

            string sSubPath;
            bool bIsExists;

            s_FolderNameWebGuiFwUpDnGrade = s_FolderNameWebGuiFwUpDnGrade.Replace("DATE", DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            sSubPath = System.Windows.Forms.Application.StartupPath + "\\report\\" + s_FolderNameWebGuiFwUpDnGrade;
            s_ReportFolderFullPath_WebGuiFwUpDnGrade = sSubPath;
            bIsExists = System.IO.Directory.Exists(sSubPath);

            if (!bIsExists)
                System.IO.Directory.CreateDirectory(sSubPath);

            s_FinalReportFilePath_WebGuiFwUpDnGrade = s_ReportPathWebGuiFwUpDnGrade + @"\" + s_FolderNameWebGuiFwUpDnGrade + @"\" + s_FinalReportFileName_WebGuiFwUpDnGrade;
            initialFinalReportExcel_WebGuiFwUpDnGrade(s_FinalReportFilePath_WebGuiFwUpDnGrade);
        }

        private void CreateNewWebGuiReport_WebGuiFwUpDnGrade()
        {
            s_WebGuiReportFileName_WebGuiFwUpDnGrade = "MODEL_WebGuiReport_SN_SW_HW_DATE.xlsx";
            s_WebGuiReportFileName_WebGuiFwUpDnGrade = getReportFileNameWebGuiFwUpDnGrade(s_WebGuiReportFileName_WebGuiFwUpDnGrade);

            s_WebGuiReportFilePath_WebGuiFwUpDnGrade = s_ReportPathWebGuiFwUpDnGrade + @"\" + s_FolderNameWebGuiFwUpDnGrade + @"\" + s_WebGuiReportFileName_WebGuiFwUpDnGrade;
            initialWebGuiReportExcel_WebGuiFwUpDnGrade(s_WebGuiReportFilePath_WebGuiFwUpDnGrade);
        }

        public string getReportFileNameWebGuiFwUpDnGrade(string reportTargetfileName)
        {
            reportTargetfileName = reportTargetfileName.Replace("MODEL", (struct_ModelStructWebGuiFwUpDnGrade.ModelName == "") ? "FwStress" : struct_ModelStructWebGuiFwUpDnGrade.ModelName);
            reportTargetfileName = reportTargetfileName.Replace("SN", struct_ModelStructWebGuiFwUpDnGrade.SN);
            reportTargetfileName = reportTargetfileName.Replace("SW", struct_ModelStructWebGuiFwUpDnGrade.SWver);
            reportTargetfileName = reportTargetfileName.Replace("HW", struct_ModelStructWebGuiFwUpDnGrade.HWver);
            reportTargetfileName = reportTargetfileName.Replace("DATE", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            return reportTargetfileName;
        }

        private void initialFinalReportExcel_WebGuiFwUpDnGrade(string filePath)
        {
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp = new Excel.Application();

            /*** Set Excel visible ***/
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Visible = true;

            /*** Do not show alert ***/
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.DisplayAlerts = false;

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.UserControl = true;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Interactive = false;

            //Set font and font size attributes
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.StandardFont = "Times New Roman";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.StandardFontSize = 11;

            /*** This method is used to open an Excel workbook by passing the file path as a parameter to this method. ***/
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkBook = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Workbooks.Add(misValue);

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkBook.Sheets[1];
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Name = "Test Results";
            createExcelTitle_WebGuiFwUpDnGrade();

            SaveAsNewExcel_Common(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade, filePath);
        }

        private void initialWebGuiReportExcel_WebGuiFwUpDnGrade(string filePath)
        {
        }

        private void createExcelTitle_WebGuiFwUpDnGrade()
        {
            //*** SettingExcelFont() ***//
            //--- [ColorIndex] 黑:1, 紅:3, 藍:5,淡黃:27, 淡橘:44 
            //--- [Font Style] U:底線 B:粗體 I:斜體 (可同時設定多個屬性)   ex:"UBI"、"BI"、"I"、"B"...

            //*** SettingExcelAlignment() ***//
            //--- [direction] H:水平方向 V:垂直方向 (可同時設定多個屬性)   ex:"HV"、"H"、"V"
            //--- [AlignmentType] C:置中 L:靠左 R:靠右


            //--------------------------------------- Report Title ---------------------------------------------//
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 2];
            SettingExcelFontUSBStorage(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "Times New Roman", 14, "U", 1);
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 2] = "CyberATE Web GUI FW Stress Final Report";

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[4, 2] = "Test";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[5, 2] = "Station";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[6, 2] = "Start Time";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[7, 2] = "End Time";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[8, 2] = "Model";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[9, 2] = "Serial";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[10, 2] = "SW Version";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[11, 2] = "HW Version";

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[4, 3] = "Web GUI FW Stress Test";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[5, 3] = "CyberTAN ATE-" + struct_ModelStructWebGuiFwUpDnGrade.ModelName;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[6, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[8, 3] = struct_ModelStructWebGuiFwUpDnGrade.ModelName;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[9, 3] = struct_ModelStructWebGuiFwUpDnGrade.SN;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[10, 3] = struct_ModelStructWebGuiFwUpDnGrade.SWver;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[11, 3] = struct_ModelStructWebGuiFwUpDnGrade.HWver;
            //--------------------------------------------------------------------------------------------------//



            //---------------------------------------- Script Info ---------------------------------------------//
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[3, 8], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[7, 9]].Borders.LineStyle = 1;  // 畫表格框線
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[3, 8] = "Option";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[3, 9] = "Value";

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[3, 8], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[3, 9]].Font.Underline = true;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[3, 8], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[3, 9]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[3, 8], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[3, 9]].Font.FontStyle = "Bold";

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range["I4"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[4, 8] = "Test Times";
            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[4, 9] = i_TestRunWebGuiFwUpDnGrade;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[5, 8] = "Final Script";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[5, 9] = s_ScriptPath_WebGuiFwUpDnGrade + @"\FwStressFinalTestItems.xlsx";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[6, 8] = "Steps Script";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[6, 9] = s_ScriptPath_WebGuiFwUpDnGrade + @"\FwStressSteps.xlsx";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[7, 8] = "WebGUI Script";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[7, 9] = struct_TestSettingWebGuiFwUpDnGrade.DeivceTestScriptPath;
            //--------------------------------------------------------------------------------------------------//



            //------------------------------------- Test Result Count ------------------------------------------//
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[10, 8], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[12, 10]].Borders.LineStyle = 1;  // 畫表格框線
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[10, 8] = "Item";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[10, 9] = "PASS";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[10, 10] = "FAIL";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[11, 8] = "FW Downgrade";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[12, 8] = "FW Upgrade";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[10, 8], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[12,10]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;  // 水平置中
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[11, 8], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[12, 8]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;    // 水平靠左
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[10, 8], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[10,10]].Font.FontStyle = "Bold";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range["H10"].Font.ColorIndex = 1;  // 黑色
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range["I10"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue); ;  // 淡藍色
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range["J10"].Font.ColorIndex = 3;  // 紅色
            //--------------------------------------------------------------------------------------------------//



            //---------------------------------------- Report Data ---------------------------------------------//
            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 2] = "Item Index";
            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 3] = "Test Item";
            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 4] = "Item Section";
            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 5] = "Test Result";
            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 6] = "Comment";
            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 7] = "Log File";

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 2] = "Test Times";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 3] = "Item Index";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 4] = "Test Item";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 5] = "Item Section";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 6] = "Test Result";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 7] = "Comment";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 8] = "Log File";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 9] = "Screen Shot";

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[14, 2], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[14, 10]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 2], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 10]].Font.FontStyle = "Bold";
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 2], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[15, 10]];
            SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "C");

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[4, 3], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[5, 3]];
            SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "L");

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[6, 3], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[11, 3]];
            SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "C");
            //--------------------------------------------------------------------------------------------------//


            /*** Set cells width ***/
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 1], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 1]].ColumnWidth = 2;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 2], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 2]].ColumnWidth = 15;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 3], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 3]].ColumnWidth = 15;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 4], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 4]].ColumnWidth = 20;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 5], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 5]].ColumnWidth = 15;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 6], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 6]].ColumnWidth = 15;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 7], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 7]].ColumnWidth = 15;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 8], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 8]].ColumnWidth = 15;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 9], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 9]].ColumnWidth = 15;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 10], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 10]].ColumnWidth = 15;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Range[st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[2, 11], st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[10, 11]].ColumnWidth = 15;


            /*** Set cells width ***/
            //SetExcelCellsWidth();

        }

        private void WriteStepTestResultToArray()
        {
            //** st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX]   <-- Current Step

            for (DEVICE_TEST_SCRIPT_STRUCT_INDEX = 0; DEVICE_TEST_SCRIPT_STRUCT_INDEX < st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade.Length; DEVICE_TEST_SCRIPT_STRUCT_INDEX++)   //------------ Deivce Test Script Data Array -----------//
            {
                string strCurrentStepAction = st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_SCRIPT_STRUCT_INDEX].Action;

                if (strCurrentStepAction == st_ReadDeviceTestScriptDataWebGuiFwUpDnGrade[DEVICE_TEST_SCRIPT_STRUCT_INDEX].Procedure)
                {
                }
            }
        }

        private void WriteFinalTestResultToReport()
        {
            int iStepsStartIndex = Convert.ToInt32(st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].StartIndex);
            int iStepsStopIndex = Convert.ToInt32(st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].StopIndex);
            
            for (int STEP_STRUCT_INDEX = 0; STEP_STRUCT_INDEX < st_ReadStepsScriptDataWebGuiFwUpDnGrade.Length; STEP_STRUCT_INDEX++)                                     //------------------ Step Script Array------------------//
            {
                int icurStepsIndex = Convert.ToInt32(st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_STRUCT_INDEX].Index);

                if (icurStepsIndex >= iStepsStartIndex && icurStepsIndex <= iStepsStopIndex)  // iStepsStartIndex <=  icurStepsIndex <= iStepsStopIndex
                {
                    string CurrentStepTestResult = st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_STRUCT_INDEX].TestResult;
                    
                    if (CurrentStepTestResult == null || CurrentStepTestResult == string.Empty || CurrentStepTestResult == "")
                    {
                        st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].TestResult = "PASS";
                    }
                    else if (CurrentStepTestResult.CompareTo("FAIL") == 0)
                    {
                        st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].TestResult = "FAIL";
                        //st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].Comment = st_ReadStepsScriptDataWebGuiFwUpDnGrade[STEP_STRUCT_INDEX].Note;
                        break;
                    }
                }
            }

            WriteFinalTestResultExcel(st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX]);
        }

        private void WriteFinalTestResultExcel(FinalTestItemsScriptDataWebGuiFwUpDnGrade FinalScriptPara)
        {
            string TEST_ITEM = st_ReadTestItemsScriptWebGuiFwUpDnGrade[TEST_ITEMS_STRUCT_INDEX].TestItem;
            string StatusCount = string.Empty;
            int gridROW;
            int gridCELL;
            int excelROW;
            int excelCELL;

            //string strCurTestTimes = string.Format("{0}/{1}", i_TestRunWebGuiFwUpDnGrade.ToString(), struct_TestSettingWebGuiFwUpDnGrade.TestRun);
            string strCurTestTimes = string.Format("{0}", i_TestRunWebGuiFwUpDnGrade.ToString());
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[4, 9] = strCurTestTimes;  //----- Update Current Test Times

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 2] = i_TestRunWebGuiFwUpDnGrade.ToString();
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 2];
            SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "C");

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 3] = FinalScriptPara.ItemIndex;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 3];
            SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "C");

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 4] = FinalScriptPara.TestItem;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 4];
            SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "L");

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 5] = FinalScriptPara.ItemSection;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 5];
            SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "L");

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 6] = FinalScriptPara.TestResult;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 6];
            SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "C");
            if (FinalScriptPara.TestResult == "PASS")
            {
                st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

                gridROW = TestStatusDict[TEST_ITEM].GridPassRow;
                gridCELL = TestStatusDict[TEST_ITEM].GridPassCell;
                excelROW = TestStatusDict[TEST_ITEM].ExcelPassRow;
                excelCELL = TestStatusDict[TEST_ITEM].ExcelPassCell;
                if (TEST_ITEM.CompareTo("FW Downgrade") == 0)
                {
                    FwDowngrade.PassCount = FwDowngrade.PassCount + 1;
                    StatusCount = FwDowngrade.PassCount.ToString();
                }
                else if (TEST_ITEM.CompareTo("FW Upgrade") == 0)
                {
                    FwUpgrade.PassCount = FwUpgrade.PassCount + 1;
                    StatusCount = FwUpgrade.PassCount.ToString();
                }
                st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[excelROW, excelCELL] = StatusCount;                                                   //----- Update Test Result to Excel Report
                Invoke(new AddDataGridViewDelegate_WebGuiFwUpDnGrade(ModifyCurrentTeatStatusToDataGridView), new object[] { gridROW, gridCELL, StatusCount });  //----- Update Test Result to GUI Data Grid View

            }
            else if (FinalScriptPara.TestResult == "FAIL")
            {
                st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                gridROW = TestStatusDict[TEST_ITEM].GridFailRow;
                gridCELL = TestStatusDict[TEST_ITEM].GridFailCell;
                excelROW = TestStatusDict[TEST_ITEM].ExcelFailRow;
                excelCELL = TestStatusDict[TEST_ITEM].ExcelFailCell;
                //----- Update Test Result to GUI Data Grid View -----//
                if (TEST_ITEM.CompareTo("FW Downgrade") == 0)
                {
                    FwDowngrade.FailCount = FwDowngrade.FailCount + 1;
                    StatusCount = FwDowngrade.FailCount.ToString();
                }
                else if (TEST_ITEM.CompareTo("FW Upgrade") == 0)
                {
                    FwUpgrade.FailCount = FwUpgrade.FailCount + 1;
                    StatusCount = FwUpgrade.FailCount.ToString();
                }
                st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelApp.Cells[excelROW, excelCELL] = StatusCount;                                                   //----- Update Test Result to Excel Report
                Invoke(new AddDataGridViewDelegate_WebGuiFwUpDnGrade(ModifyCurrentTeatStatusToDataGridView), new object[] { gridROW, gridCELL, StatusCount });  //----- Update Test Result to GUI Data Grid View


                //--- When Test Fail, Add Screen Shot File Link
                //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 9] = FinalScriptPara.ScreenShotPath;
                st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 9];
                st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "C");
                st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Hyperlinks.Add(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, FinalScriptPara.ScreenShotPath, Type.Missing, FinalScriptPara.ScreenShotPath, "ScreenShot");
            }
            else if (FinalScriptPara.TestResult == "ERROR")
            {
                st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }

            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 7] = FinalScriptPara.Comment;
            st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 7];
            SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "C");


            ////Log File Link
            ////st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 8] = FinalScriptPara.Log;
            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange = st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Cells[i_DATA_FinalROW_WebGuiFwUpDnGrade, 8];
            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            //SettingExcelAlignment(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "H", "C");
            //st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelWorkSheet.Hyperlinks.Add(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, FinalScriptPara.ScreenShotPath, Type.Missing, FinalScriptPara.Log, "ConsoleLog");


            //SettingExcelFontUSBStorage(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade.excelRange, "Times New Roman", 11, "B", 3);
            SaveExcelFile_Common(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade);  // Save Excel File (每項測完都存檔一次，避免當機檔案遺失)
            i_DATA_FinalROW_WebGuiFwUpDnGrade++;
        }

        #endregion


        //=================================================================//
        //------------------------ Invoke Function ------------------------//
        //=================================================================//
        #region Invoke Function

        private void ClearInfoTextBoxWebGuiFwUpDnGrade()
        {
            txtWebGuiFwUpDnGradeFunctionTestInformation.Clear();
        }
        
        private void ToggleWebGuiFwUpDnGradeTestTimes()
        {
            labelWebGuiFwUpDnGradeTestTimes.Text = i_TestRunWebGuiFwUpDnGrade.ToString();
        }

        private void ModifyCurrentTeatStatusToDataGridView(int iROW, int iCELL, string dataValue)
        {
            dgvWebGuiFwUpDnGradeFunctionTestDataView.Rows[iROW].Cells[iCELL].Value = dataValue;
        }


        //-----------------  System Cursor ------------------//
        #region System Cursor
        private void ToggleWebGuiFwUpDnGradeFunctionTestSystemCursorWait()
        {
            ToggleWebGuiFwUpDnGradeFunctionTestSystemCursorStatus(true);
        }

        private void ToggleWebGuiFwUpDnGradeFunctionTestSystemCursorDefault()
        {
            ToggleWebGuiFwUpDnGradeFunctionTestSystemCursorStatus(false);
        }

        private void ToggleWebGuiFwUpDnGradeFunctionTestSystemCursorStatus(bool Toggle)
        {
            if (Toggle == true)
            {
                btnWebGuiFwUpDnGradeFunctionTestRun.Enabled = false;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            }
            else
            {
                btnWebGuiFwUpDnGradeFunctionTestRun.Enabled = true;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
        }
        #endregion
        //---------------------------------------------------//



        //------------------ Test Controller ----------------//
        #region Test Controller
        private void ToggleWebGuiFwUpDnGradeFunctionTestSettingTrue()
        {
            ToggleWebGuiFwUpDnGradeFunctionTestController(true);
            //Debug.WriteLine("Toggle");
        }

        private void ToggleWebGuiFwUpDnGradeFunctionTestController(bool Toggle)
        {
            btnWebGuiFwUpDnGradeFunctionTestRun.Text = Toggle ? "Run" : "Stop";


            //-------- MenuItem -------//
            fileToolStripMenuItem.Enabled = Toggle;
            itemToolStripMenuItem.Enabled = Toggle;
            setupToolStripMenuItem.Enabled = Toggle;
            helpToolStripMenuItem.Enabled = Toggle;

            //----- Function Test -----//
            gbox_ModelWebGuiFwUpDnGrade.Enabled = Toggle;
            gbox_DeviceSettingWebGuiFwUpDnGrade.Enabled = Toggle;
            gbox_TestSettingWebGuiFwUpDnGrade.Enabled = Toggle;


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


            if (bWebGuiFwUpDnGradeFTThreadRunning == false)
            {
                btnWebGuiFwUpDnGradeFunctionTestRun.Enabled = false;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

                try
                {
                    if (threadWebGuiFwUpDnGradeFTstopEvent.IsAlive)
                    {
                        threadWebGuiFwUpDnGradeFTstopEvent.Abort();
                        threadWebGuiFwUpDnGradeFTstopEvent.Join();
                    }
                    if (threadWebGuiFwUpDnGradeFTscriptItem.IsAlive)
                    {
                        threadWebGuiFwUpDnGradeFTscriptItem.Abort();
                        threadWebGuiFwUpDnGradeFTscriptItem.Join();
                    }
                    if (threadWebGuiFwUpDnGradeFT.IsAlive)
                    {
                        threadWebGuiFwUpDnGradeFT.Abort();
                        threadWebGuiFwUpDnGradeFT.Join();
                    }
                }
                catch { }

                TestFinishedActionWebGuiFwUpDnGrade();


                MessageBoxTopMost("ATE Information", "Test complete !!!");  // 將 MessageBox於桌面置頂
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                btnWebGuiFwUpDnGradeFunctionTestRun.Enabled = true;
                bWebGuiFwUpDnGradeTestComplete = true;
            }
        }
        #endregion
        //---------------------------------------------------//
        
        #endregion
        
#endregion

    }


}
