using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;
using System.Net;
using System.Net.NetworkInformation;
using System.IO;
using AgilentInstruments;
using ComportClass;
using NS_CbtSeleniumApi;



namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {       
        /* Declare delegate prototype */
        private delegate void SetTextCallBackT(string text, TextBox textbox);
        public delegate void showRouterGUIDelegate();
        /* Declare delegate function */
        private delegate void showWirelessPatameterContentDelegate(WirelessParameter wp, TextBox textbox);
        
        //Thread threadRouterMainFT;
        //bool bRouterMainFTThreadRunning = false;
        
        /* Declare global variable for all test items */
        /* For RvR series*/
        string Path_runtst;
        string Path_fmttst;
        Agilent11713A[] a11713A_2_4G;
        Agilent11713A[] a11713A_5G;
        decimal[] Attenuation_buttonValue_2_4G;
        decimal[] Attenuation_buttonValue_5G;        
        bool bConfigRouter = true;
        int Channel_2_4G_Bound = 11;
        int[] Channel_5G_20M = { 36, 40, 44, 48, 149, 153, 157, 161, 165 };
        int[] Channel_5G_40M = { 36, 40, 44, 48, 149, 153, 157, 161 };
        int[] Channel_5G_80M = { 36, 40, 44, 48, 149, 153, 157, 161 };
        string[] Mode_2_4G = { "11N_20M", "11N_40M", "HT40+", "HT40-"};
        string[] Mode_5G = { "11N_20M", "11N_40M", "11AC_20M", "11AC_40M", "11AC_80M", "HT40+", "HT40-"};

        struct WirelessParameter
        {
            public string band;
            public string mode;
            public string ssid_config;
            public string ssid_text;
            public string channel_config;
            public int channel_start;
            public int channel_stop;
            public string security_config;
            public string security_mode;
            public string key_index;
            public string passphrase;            
        }

        int PingTimeout = 1000;

        string model;
        string security = "None";
        string passphrase = string.Empty;

        Thread threadRouterFT;
        bool bRouterFTThreadRunning = false;

        public string DutConfigurationType = string.Empty;
        public string DutConfigurationFile = string.Empty;        

        /* Excel object */
        Excel.Application excelAppRouter;
        Excel.Workbook excelWorkBookRouter;
        Excel.Worksheet excelWorkSheetRouter;
        Excel.Range excelRangeRouter;

        object misValue = System.Reflection.Missing.Value;

        public class WifiBasic
        {
            public string band;
            public string mode;
            public int channel;
            public string ssid;
            public string security;
            public string passphrase;
        }

        public class AtteuationValue
        {
            public int start;
            public int stop;
            public int steps;
        }

        public class TstFile
        {
            public string txTst;
            public string rxTst;
            public string biTst;
        }

        struct RouterDutsSetting
        {
            public int index;
            public string IpAddress;
            public string MacAddress;
            public string PcIpAddress;
            public string SwitchPort;
            public string ComportNum;
            public string ModelName;
            public string SwVersion;
            public string HwVersion;
            public string SerialNumber;
            public string GuiScriptExcelFile;
        }

        struct WifiParameter
        {
            public string SSID;
            public string Channel;
            public string Band;
            public string Mode;
            public string BandWidth;
            public string Security;
            public string SecurityMode;
            public string SecurityKey;
        }

        struct RouterIntegrationTestCase
        {
            public int iIndex;
            public string TestName;
            public string TestFunction;


        }

        string[,] sa_RounterTestCondition;
        string[,] sa_RounterStepsCondition;
        string[,] sa_RounterDutsSetting;
        string[,] sa_RounterGuiScript;

        RouterDutsSetting[] st_DutsSetting;
        RouterDutsSetting st_CurrentDut;
        //st_CurrentTestCase ;
        WifiParameter st_WifiParameter;


        //------------------------------------------------//
        //----------- Router Test Script struct ----------//
        //------------------------------------------------//
        #region Router Test struct
        struct FinalTestItemsScriptData_RouterTest
        {
            public string ItemIndex;
            public string TestItem;
            public string ItemSection;
            public string StartIndex;
            public string StopIndex;
            public string TestResult;
            public string Comment;
            public string Log;
            public string ScreenShotPath;
        }

        struct StepsScriptData_RouterTest
        {
            public string Index;
            public string FunctionName;
            public string Name;
            public string Action;
            public string TestConditionDUTindex;
            public string[] Command;
            public string GetDataCondition;
            public string RemoteSendString;
            //public string WriteValue;
            public string ExpectedValue;
            public string GetValue;
            public string TestTimeOut;
            public string TestResult;
            public string InfoLogPath;
            public string[] FileAndScreenshotPath;
            public string Note;
        }

        struct DeviceWebGuiScriptData_RouterTest
        {
            public string Procedure;
            public string TestIndex;
            public string Action;
            public string ActionName;
            public string ElementType;
            public string RadioButtonExpectedValueXpath;
            public string WriteExpectedValue;
            public string ElementXpath;
            public string TestTimeOut;
            public string GetValue;
        }
        #endregion


        /*============================================================================================*/
        /*================================= Excel Function Area   ====================================*/
        /*============================================================================================*/
        #region Excel Function Area

        private void SaveAsNewExcel_Common(ExcelObject excelObject, string savePath)
        {
            try
            {
                excelObject.excelWorkBook.SaveAs(savePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Error: Check the file path must within 218 characters.");
            }
        }

        private void SaveExcelFile_Common(ExcelObject excelObject)
        {
            try
            {
                if (excelObject.excelWorkBook != null)
                {
                    /* Save excel data */
                    excelObject.excelWorkBook.Save();
                }
            }
            catch { }
        }

        private void CloseExcel_Common(ExcelObject excelObject)
        {
            try
            {
                /* Turn on interactive mode */
                excelObject.excelApp.Interactive = true;
                excelObject.excelWorkBook.Close();
                excelObject.excelApp.Quit();

                /*System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppRfi);*/
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelObject.excelWorkSheet);
                excelObject.excelWorkSheet = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelObject.excelWorkBook);
                excelObject.excelWorkBook = null;
                releaseObject_Common(excelObject.excelRange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelObject.excelRange);
                excelRangeCommon = null;
                releaseObject_Common(excelObject.excelApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelObject.excelApp);
                excelObject.excelApp = null;
                GC.Collect();
            }
            catch (Exception ex)
            { }
        }

        private void Save_and_Close_Excel_File_Common(ExcelObject excelObject)
        {
            try
            {
                if (excelObject.excelWorkBook != null)
                {
                    /* Save excel data */
                    excelObject.excelWorkBook.Save();
                    CloseExcel_Common(excelObject);
                }
            }
            catch { }
        }

        private void SaveEndTimetoExcel_Common(ExcelObject excelObject, int sheetNum, int CELL1, int CELL2)
        {
            try
            {
                excelObject.excelWorkSheet = excelObject.excelWorkBook.Sheets[1];
                excelObject.excelApp.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");  // Save End Time to Excel
            }
            catch { }
        }
        
        private void SetExcelCellsHigh_Common(ExcelObject excelObject, string CELL1, string CELL2)
        {
            excelObject.excelWorkSheet = excelObject.excelWorkBook.Sheets[1];
            excelObject.excelRange = excelObject.excelWorkSheet.Range[CELL1, CELL2];
            excelObject.excelRange.RowHeight = 15;
        }
        
        private void saveExcelRouter(string savePath)
        {
            try
            {
                excelWorkBookRouter.SaveAs(savePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Error: Total of file path is exceeds 218 characters.");
            }
        }

        private void Save_and_Close_Excel_File()
        {
            try
            {
                if (excelWorkBookCommon != null)
                {
                    /* Save excel data */
                    excelWorkBookCommon.Save();
                    closeExcelCommon();
                }
            }
            catch { }

        }

        private void closeExcelCommon()
        {
            /* Turn on interactive mode */
            excelAppCommon.Interactive = true;
            excelWorkBookCommon.Close();
            excelAppCommon.Quit();

            /*System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppRfi);*/
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkSheetCommon);
            excelWorkSheetCommon = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkBookCommon);
            excelWorkBookCommon = null;
            releaseObject_Common(excelRangeCommon);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelRangeCommon);
            excelRangeCommon = null;
            releaseObject_Common(excelAppCommon);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppCommon);
            excelAppCommon = null;
            GC.Collect();
        }

        private void SaveEndTime_and_IEversion_to_Excel(CBT_SeleniumApi csUSBStorageBrowser)
        {
            try
            {
                excelWorkSheetCommon = excelWorkBookCommon.Sheets[1];

                excelAppCommon.Cells[10, 4] = "Script Path";
                excelAppCommon.Cells[10, 5] = txtUSBStorageFunctionTestScriptFilePath.Text;
                excelRangeCommon = excelWorkSheetCommon.Cells[10, 5];
                SettingExcelAlignment(excelRangeCommon, "H", "R");
                excelAppCommon.Cells[11, 4] = "Browser Version";


                excelAppCommon.Cells[11, 5] = csUSBStorageBrowser.BrowserVersion();

                excelRangeCommon = excelWorkSheetCommon.Cells[11, 5];
                SettingExcelAlignment(excelRangeCommon, "H", "R");



                excelAppCommon.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");  // Save End Time to Excel
            }
            catch { }

        }       

        private void closeExcelRouter()
        {
            /* Turn on interactive mode */
            excelAppRouter.Interactive = true;
            excelWorkBookRouter.Close();
            excelAppRouter.Quit();

            /*System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppRfi);*/
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkSheetRouter);
            excelWorkSheetRouter = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkBookRouter);
            excelWorkBookRouter = null;
            releaseObject_Router(excelRangeRouter);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelRangeRouter);
            excelRangeRouter = null;
            releaseObject_Router(excelAppRouter);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppRouter);
            excelAppRouter = null;
            GC.Collect();
        }

        private void releaseObject_Router(object obj)
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

        private void SettingExcelAlignment(Excel.Range excelRangeCGA2121, string direction, string AlignmentType)
        {
            //--- [direction] H:水平方向 V:垂直方向 (可同時設定多個屬性)   ex:"HV"、"H"、"V"
            //--- [AlignmentType] C:置中 L:靠左 R:靠右

            if (AlignmentType.CompareTo("C") == 0)      // 置中
            {
                if (direction.ToLower().IndexOf("h") >= 0)
                    excelRangeCGA2121.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                if (direction.ToLower().IndexOf("v") >= 0)
                    excelRangeCGA2121.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            else if (AlignmentType.CompareTo("L") == 0) // 靠左
            {
                if (direction.ToLower().IndexOf("h") >= 0)
                    excelRangeCGA2121.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                if (direction.ToLower().IndexOf("v") >= 0)
                    excelRangeCGA2121.VerticalAlignment = Excel.XlHAlign.xlHAlignLeft;
            }
            else if (AlignmentType.CompareTo("R") == 0) // 靠右
            {
                if (direction.ToLower().IndexOf("h") >= 0)
                    excelRangeCGA2121.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                if (direction.ToLower().IndexOf("v") >= 0)
                    excelRangeCGA2121.VerticalAlignment = Excel.XlHAlign.xlHAlignRight;
            }
        }

        private void SettingExcelFont_Common(Excel.Range excelRangeCommon, string FONT_TYPE, int FONT_SIZE, string FONT_STYLE, int COLOR_INDEX)
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

        #endregion //Excel Function Area



        private string createRouterSubFolder(string ModelName)
        {
            return ((ModelName == "") ? "Router_" : ModelName + "_") + DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        /* delegate call back function */
        private void SetText(string msg, TextBox txtInformation)
        {
            txtInformation.AppendText(msg + Environment.NewLine);
        }

        private void SetTextC(string msg, TextBox txtInformation)
        {
            txtInformation.Clear();
            txtInformation.AppendText(msg);
        }

        private bool CsvAppend(string filePath, string str)
        {
            if (!File.Exists(filePath))
            {
                Debug.WriteLine("File doesn't exist");
                return false;
            }

            try
            {
                StringBuilder csv = new StringBuilder();
                csv.AppendLine(str);
                File.AppendAllText(filePath, csv.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Write CSV Failed");
                return false;   
            }

            return true;
        }

        private bool InserLine2TextFile(string filePath, int iInsertLine, string text)
        {
            if (!File.Exists(filePath)) return false;
            //string sTestFileName = @"e:\t1.txt";
            //int iInsertLine = 5;
            //string text = "插入的内容";
            string sText = "";
            System.IO.StreamReader sr = new System.IO.StreamReader(filePath);

            int iLnTmp = 0; //记录文件行数
            while (!sr.EndOfStream)
            {
                iLnTmp++;
                if (iLnTmp == iInsertLine)
                {
                    sText += text + "\r\n";  //将值插入
                }
                string sTmp = sr.ReadLine();    //记录当前行
                sText += sTmp + "\r\n";
            }
            sr.Close();

            System.IO.StreamWriter sw = new System.IO.StreamWriter(filePath, false);
            sw.Write(sText);
            sw.Flush();
            sw.Close();
            return true;
        }

        private string ThroughputValue(string filePath)
        {
            if (!File.Exists(filePath)) return null;
                        
            string sText = "";
            System.IO.StreamReader sr = new System.IO.StreamReader(filePath);

            while (!sr.EndOfStream)
            {
                string sTmp = sr.ReadLine();    //记录当前行
                sText = sTmp;
            }
            sr.Close();

            string[] t = sText.Split(',');
            return t[9];
        }

        private string ThroughputUnit(string filePath)
        {
            if (!File.Exists(filePath)) return null;

            string sText = "";
            string preText = "";
            System.IO.StreamReader sr = new System.IO.StreamReader(filePath);

            while (!sr.EndOfStream)
            {
                string sTmp = sr.ReadLine();    //记录当前行
                preText = sText;
                sText = sTmp;
            }
            sr.Close();

            string[] t = preText.Split(',');
            return t[9];
        }

        private bool ConfigureRouter(string modelName, string hostIP, string userName, string passWord, WifiBasic wifi)
        {
            switch (modelName.ToLower())
            {
                case "tch-2.4g":
                    return Tch24GDeviceConfigure(hostIP, userName, passWord, wifi);                    
                case "tch-dual":
                    return Tch5GDeviceConfigure(hostIP, userName, passWord, wifi);
                default:
                    return false;                    
            }            
        }

        private bool WebHttpGetData(string host, string agent, bool keepAlive, bool Auth, string username, string password, string compareText, ref string response, string contentType = "")
        {
            string _uri = host;
            string _strHtml = string.Empty;
            
            try
            {
                WebRequest myRequest = WebRequest.Create(_uri);
                myRequest.Timeout = 30000;
                if (Auth)
                    myRequest.Credentials = new NetworkCredential(username, password);

                HttpWebRequest myHttpWebRequest = (HttpWebRequest)myRequest;
                myHttpWebRequest.Timeout = 30000;
                myHttpWebRequest.Method = "GET";
                myHttpWebRequest.UserAgent = agent;
                myHttpWebRequest.KeepAlive = keepAlive;

                if (contentType != "")
                    myHttpWebRequest.ContentType = contentType;

                using (WebResponse myWebResponse = myHttpWebRequest.GetResponse())
                {
                    using (Stream myStream = myWebResponse.GetResponseStream())
                    {
                        using (StreamReader myReader = new StreamReader(myStream))
                        {
                            _strHtml = myReader.ReadToEnd();
                            //document.Load(myReader.ReadToEnd());
                        }
                    }
                    myWebResponse.Close();
                }

                //textBox1.Text = _strHtml;
                response = _strHtml;
                //response = document.DocumentNode.InnerText;

                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
        }

        private bool WebHttpPostData(string host, string data, string agent, string ContentType, bool keepAlive, string Type, bool Auth, string username, string password, ref string response)
        {
            byte[] bs = Encoding.ASCII.GetBytes(data);

            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(host);
            req.Method = "POST";
            req.ContentType = ContentType;
            req.KeepAlive = keepAlive;
            //req.ProtocolVersion = HttpVersion.Version11;
            req.UserAgent = agent;

            req.ContentLength = bs.Length;
            if (Type == "jnap")
            {
                //req.Headers.Add("
                //req.Headers.Add("x-jnap-authorization", "Basic YWRtaW46YWRtaW4=\r\n");
                //req.Headers.Add("x-jnap-action", "http://linksys.com/jnap/core/Transaction\r\n");
                //req.Headers.Add("x-requested-with", "XMLHttpRequest\r\n");
            }

            if (Auth)
            {
                req.Headers.Add("Authorization", "Basic YWRtaW46YWRtaW4=\r\n");
                req.Headers.Add("Credentials", username + ":" + password);
                //req.Headers.Add("Credentials", "admin:admin");
            }

            try
            {
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(bs, 0, bs.Length);
                }
                using (WebResponse wr = req.GetResponse())
                {
                    Stream myStream = wr.GetResponseStream();
                    StreamReader reader = new StreamReader(myStream);
                    string strHtml = reader.ReadToEnd();
                    //reader
                    //textBox1.AppendText(strHtml);
                    response = strHtml;

                    //textBox1.AppendText(Environment.NewLine);
                    //textBox1.AppendText(Environment.NewLine);


                    if (strHtml.ToLower().IndexOf("error") >= 0)
                    {
                        //textBox1.AppendText("Perfect");
                        return false;
                    }
                    wr.Close();
                    return true;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        
        /*============================================================================================*/
        /*================================  Delegate function ========================================*/
        /*============================================================================================*/
        private void showWirelessPatameterAllContent(WirelessParameter wp, TextBox textbox)
        {
            string str = string.Empty;
            str += "band: " + wp.band;
            str += "\tmode: " + wp.mode;
            str += "\tssid_config: " + wp.ssid_config;
            str+= "\tssid_text: " + wp.ssid_text ;
            str += "\tchannel_config: " + wp.channel_config;
            str += "\tchannel_start: " + wp.channel_start;
            str += "\tchannel_stop: " + wp.channel_stop;
            str += "\tsecurity_config: " + wp.security_config;
            str += "\tsecurity_mode: " + wp.security_mode;
            str += "\tkey_index: " + wp.key_index;
            str += "\tpassphrase: " + wp.passphrase;         

            textbox.AppendText(str + Environment.NewLine);
        }

        private void showWirelessPatameterPartContent(WirelessParameter wp, TextBox textbox)
        {
            string str = string.Empty;
            str += "band: " + wp.band;
            str += "\tmode: " + wp.mode;
            str += "\tssid_config: " + wp.ssid_config;
            str += "\tssid_text: " + wp.ssid_text;
            str += "\tchannel_config: " + wp.channel_config;
            str += "\tchannel_start: " + wp.channel_start;
            str += "\tchannel_stop: " + wp.channel_stop;
            str += "\tsecurity_config: " + wp.security_config;
            str += "\tsecurity_mode: " + wp.security_mode;
            str += "\tkey_index: " + wp.key_index;
            str += "\tpassphrase: " + wp.passphrase;

            textbox.AppendText(str + Environment.NewLine);
        }


        /*============================================================================================*/
        /*============================= Print Screen Function Area   =================================*/
        /*============================================================================================*/
        #region Print Screen Function Area

        private string PrintFullScreen_Common(string FilePath)
        {
            string strInfo = string.Empty;

            try
            {
                Bitmap myImage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                Graphics g = Graphics.FromImage(myImage);
                g.CopyFromScreen(new Point(0, 0), new Point(0, 0), new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
                IntPtr dc1 = g.GetHdc();
                g.ReleaseHdc(dc1);
                myImage.Save(FilePath);
            }
            catch (Exception info)
            {
                strInfo = ("PrintFullScreen_Common() Error: " + info.ToString());
            }
            return strInfo;
        }

        private string PrintFocusedWindow_Common(string FilePath)
        {
            string strInfo = string.Empty;

            try
            {
                Bitmap myImage = new Bitmap(this.Width, this.Height);
                Graphics g = Graphics.FromImage(myImage);
                g.CopyFromScreen(new Point(this.Location.X, this.Location.Y), new Point(0, 0), new Size(this.Width, this.Height));
                IntPtr dc1 = g.GetHdc();
                g.ReleaseHdc(dc1);
                myImage.Save(FilePath);
            }
            catch (Exception info)
            {
                strInfo = ("PrintFocusedWindow_Common Error: " + info.ToString());
            }
            return strInfo;
        }
        
        #endregion


        /*============================================================================================*/
        /*=============================== GUI Test Function Area   ===================================*/
        /*============================================================================================*/
        #region GUI Test Function Area

        /* Declare delegate prototype */
        //public delegate void showHddtaRfiRestartDelegate(string str);
        //public delegate void showHddtaRfiChannelPowerDelegate(string ch_power, int frequency);
        //public delegate void showHddtaRfiResultDelegate(int CenterFrequency, double Level, string Constellation, string ReportPwr, string Snr, string CE, string UE);

        public delegate void showCommonGUIDelegate();

        /* Declare delegate prototype */
        //private delegate void SetTextCallBack(string text, TextBox textbox);


        struct FinalScriptDataGUI
        {
            public string TestItem;
            public string FunctionName;
            public string Name;
            public string StartIndex;
            public string StopIndex;
            public string TestResult;
            public string Comment;
            public string Log;
        }

        struct StepsScriptDataGUI
        {
            public string Index;
            public string FunctionName;
            public string Name;
            public string Action;
            public string TestConditionDUTindex;
            public string[] Command;
            public string GetDataCondition;
            public string[] WriteValue;
            public string ExpectedValue;
            public string GetValue;
            public string TestTimeOut;
            public string TestResult;
            public string InfoLogPath;
            public string[] FileAndScreenshotPath;
            public string Note;
            public string FtpUsername;
            public string FtpPassword;
            public string FtpIP;
            public string FtpServerFolder;
            public string FtpServerPort;
            public string FtpRenameTo;
            public string FtpPassiveMode;
        }

        struct ScriptDataGUI
        {
            public string Procedure;
            public string TestIndex;
            public string Action;
            public string ActionName;
            public string ElementType;
            public string RadioButtonExpectedValueXpath;
            public string WriteExpectedValue;
            public string ElementXpath;
            public string TestTimeOut;
            public string GetValue;
        }

        public struct ModelGroupStruct
        {
            public string ModelName;
            public string SN;
            public string SWver;
            public string HWver;

            public ModelGroupStruct(string ModelName, string SN, string SWver, string HWver)
            {
                this.ModelName = ModelName;
                this.SN = SN;
                this.SWver = SWver;
                this.HWver = HWver;
            }
        }

        struct TestBrowerSetting
        {
            public string TestRun;
            public bool StopWhenTestError;
        }

        struct LoginSetting
        {
            public string GatewayIP;
            public string UserName;
            public string Password;
            public string HTTP_Port;
            public int RebootWaitTime;
            public bool ConsoleLog;
        }

        private string ReadConsoleLogByBytesGUI(Comport com)
        {
            int byteToRead = com.GetBytesToRead();
            return com.Read(0, byteToRead);
        }

        private void saveExcelGUI(string savePath)
        {
            try
            {
                st_ExcelObjectGUI.excelWorkBook.SaveAs(savePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Error: Check the file path must within 218 characters.");
            }
        }

        private void saveWebGUIExcelGUI(string savePath)
        {
            try
            {
                excelWorkBookCommon.SaveAs(savePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Error: Check the file path must within 218 characters.");
            }
        }


        //------------------------------------------------//
        //-------- Web GUI FW Up/Downgrade struct --------//
        //------------------------------------------------//

        #region Web GUI FW Up/Downgrade struct

        struct CurrentTestStatus
        {
            public int PassCount;
            public int FailCount;
            public int GridPassRow;
            public int GridPassCell;
            public int GridFailRow;
            public int GridFailCell;
            public int ExcelPassRow;
            public int ExcelPassCell;
            public int ExcelFailRow;
            public int ExcelFailCell;
        }

        struct ModelStructWebGuiFwUpDnGrade
        {
            public string ModelName;
            public string SN;
            public string SWver;
            public string HWver;
        }

        struct DeviceSettingWebGuiFwUpDnGrade
        {
            public string DeviceIP;
            public string HTTP_Port;
            public bool ConsoleLog;
        }

        struct TestSettingWebGuiFwUpDnGrade
        {
            public string DowngradeFwFilePath;
            public string UpgradeFwFilePath;
            public string DeivceTestScriptPath;
            public string TestRun;
            public string DefaultTimeout;
            public bool StopWhenTestError;
        }

        struct FinalTestItemsScriptDataWebGuiFwUpDnGrade
        {
            public string ItemIndex;
            public string TestItem;
            public string ItemSection;
            public string StartIndex;
            public string StopIndex;
            public string TestResult;
            public string Comment;
            public string Log;
            public string ScreenShotPath;
        }

        struct StepsScriptDataWebGuiFwUpDnGrade
        {
            public string Index;
            public string FunctionName;
            public string Name;
            public string Action;
            public string TestConditionDUTindex;
            public string[] Command;
            public string GetDataCondition;
            //public string WriteValue;
            public string ExpectedValue;
            public string GetValue;
            public string TestTimeOut;
            public string TestResult;
            public string InfoLogPath;
            public string[] FileAndScreenshotPath;
            public string Note;
        }

        struct DeviceTestScriptDataWebGuiFwUpDnGrade
        {
            public string Procedure;
            public string TestIndex;
            public string Action;
            public string ActionName;
            public string ElementType;
            public string RadioButtonExpectedValueXpath;
            public string WriteExpectedValue;
            public string ElementXpath;
            public string TestTimeOut;
            public string GetValue;
        }

        #endregion




        //------------------------------------------------//
        //------------- Guest Network struct -------------//
        //------------------------------------------------//
        #region Guest Network struct

        struct ModelStructGuestNetwork
        {
            public string ModelName;
            public string SN;
            public string SWver;
            public string HWver;
        }

        struct DeviceSettingGuestNetwork
        {
            public string DeviceIP;
            public string HTTP_Port;
            public string RebootWaitTime;
            public bool ConsoleLog;
        }

        struct GAsettingGuestNetwork
        {
            public string RemoteServerIP;
            public string RemoteServerPort;
        }
        
        struct TestSettingGuestNetwork
        {
            public string DeivceTestScriptPath;
            public string TestRun;
            public string DefaultTimeout;
            public bool StopWhenTestError;
        }

        #endregion


        



        //struct ScriptData
        //{
        //    public string TestIndex;
        //    public string TestStep;
        //    public string Action;
        //    public string ActionName;
        //    public string ElementType;
        //    public string WriteValue;
        //    public string ExpectedValue;
        //    public string RadioButtonExpectedValueXpath;
        //    public string URL;
        //    public string ElementXpath;
        //    public string ApplyButtonXpath;
        //    public string TestTimeOut;
        //    public string TriggerReboot;
        //    public string GetValue;
        //}

        private void closeExcelReadFileCommon()
        {
            /* Turn on interactive mode */
            excelReadAppCommon.Interactive = true;
            excelReadWorkBookCommon.Close(false, misValue, misValue);
            excelReadAppCommon.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelReadWorkSheetCommon);
            excelReadWorkSheetCommon = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelReadWorkBookCommon);
            excelReadWorkBookCommon = null;
            releaseObject_Common(excelReadRangeCommon);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelReadRangeCommon);
            excelReadRangeCommon = null;
            releaseObject_Common(excelReadAppCommon);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelReadAppCommon);
            excelReadAppCommon = null;
            GC.Collect();
        }

        private void releaseObject_Common(object obj)
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
        /*============================== USB Storage Function Area ===================================*/
        /*============================================================================================*/
        #region USB Storage

        struct FinalScriptDataUSBStorage
        {
            public string TestItem;
            public string FunctionName;
            public string Name;
            public string StartIndex;
            public string StopIndex;
            public string TestResult;
            public string Comment;
            public string Log;
        }

        struct StepsScriptDataUSBStorage
        {
            public string Index;
            public string FunctionName;
            public string Name;
            public string Action;
            public string TestConditionDUTindex;
            public string[] Command;
            public string GetDataCondition;
            public string[] WriteValue;
            public string ExpectedValue;
            public string GetValue;
            public string TestTimeOut;
            public string TestResult;
            public string InfoLogPath;
            public string[] FileAndScreenshotPath;
            public string Note;
            public string FtpUsername;
            public string FtpPassword;
            public string FtpIP;
            public string FtpServerFolder;
            public string FtpServerPort;
            public string FtpRenameTo;
            public string FtpPassiveMode;
        }

        struct ScriptDataUSBStorage
        {
            public string Procedure;
            public string TestIndex;
            public string Action;
            public string ActionName;
            public string ElementType;
            public string RadioButtonExpectedValueXpath;
            public string WriteExpectedValue;
            public string ElementXpath;
            public string TestTimeOut;
            public string GetValue;
        }

        //public delegate void showCommonGUIDelegate();

        Thread threadCommonFT;
        Thread threadCommonFTstopEvent;
        Thread threadCommonFTscriptItem;
        bool bCommonFTThreadRunning = false;
        bool bCommonTestComplete = true;
        bool bCommonSingleScriptItemRunning = false;


        /* Excel object */
        Excel.Application excelAppCommon = null;
        Excel.Workbook excelWorkBookCommon = null;
        Excel.Worksheet excelWorkSheetCommon = null;
        Excel.Range excelRangeCommon = null;

        Excel.Application excelReadAppCommon = null;
        Excel.Workbook excelReadWorkBookCommon = null;
        Excel.Worksheet excelReadWorkSheetCommon = null;
        Excel.Range excelReadRangeCommon = null;

        public struct ExcelObject
        {
            public Excel.Application excelApp;
            public Excel.Workbook excelWorkBook;
            public Excel.Worksheet excelWorkSheet;
            public Excel.Range excelRange;
        }

        private string ReadConsoleLogByBytesUSBStorage(Comport com)
        {
            int byteToRead = com.GetBytesToRead();
            return com.Read(0, byteToRead);
        }

        private void saveExcelUSBStorage(string savePath)
        {
            try
            {
                st_ExcelObjectUSBStorage.excelWorkBook.SaveAs(savePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Error: Check the file path must within 218 characters.");
            }
        }

        private void saveWebGUIExcelUSBStorage(string savePath)
        {
            try
            {
                excelWorkBookCommon.SaveAs(savePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Error: Check the file path must within 218 characters.");
            }
        }


        #endregion


        
        /*============================================================================================*/
        /*=================================== New Added Function =====================================*/
        /*============================================================================================*/

        public void MessageBoxTopMost(string Title, string Info) // 將 MessageBox訊息置頂
        {
            MessageBox.Show(new Form { TopMost = true }, Info, Title, MessageBoxButtons.OK);
        }

        private void HideAllTabPage()
        {
            //this.tpWebGUI_TestCondition.Parent = null;           // hide tpWebGUI_TestCondition TabPage
            //this.tpWebGUI_FunctionTest.Parent = null;            // hide tpWebGUI_FunctionTest TabPage
            this.tpGuestNetworkFunctionTest.Parent = null;
            this.tpUSBStorageFunctionTest.Parent = null;         // hide tpUSBStorage_FunctionTest TabPage
            this.tpWebGuiFwUpDnGradeFunctionTest.Parent = null;  // hide tpWebGuiFwUpDnGradeFunctionTest TabPage
            this.SallyTestPage.Parent = null;                    // hide SallyTestPage TabPage
            this.tpGuestNetworkFunctionTest.Parent = null;       // hide tpGuestNetworkFunctionTest TabPage
        }

    
    }
}



/*
 * https://dotblogs.com.tw/henry/2011/11/27/59666
 在string.Format參數中，大括號{}是有特殊意義的符號，但是如果我們希望最終的結果中包含大括號{}，那麼我們需要怎麼做呢？

是”\{”嗎？很遺憾，運行時，會給你一個Exception的！

正確的寫法是{{和}}。
雙重 {{  或 }} 即可輸出 { 或 }.

 
 
 */