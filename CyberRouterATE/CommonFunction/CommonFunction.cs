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
using System.Xml;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;


namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {       
        ///* Declare delegate prototype */
        //private delegate void SetTextCallBackT(string text, TextBox textbox);
        //public delegate void showRouterGUIDelegate();
        ///* Declare delegate function */
        //private delegate void showWirelessPatameterContentDelegate(WirelessParameter wp, TextBox textbox);
        
        ////Thread threadRouterMainFT;
        ////bool bRouterMainFTThreadRunning = false;
        
        ///* Declare global variable for all test items */
        ///* For RvR series*/
        //string Path_runtst;
        //string Path_fmttst;
        //Agilent11713A[] a11713A_2_4G;
        //Agilent11713A[] a11713A_5G;
        //decimal[] Attenuation_buttonValue_2_4G;
        //decimal[] Attenuation_buttonValue_5G;        
        //bool bConfigRouter = true;
        //int Channel_2_4G_Bound = 11;
        //int[] Channel_5G_20M = { 36, 40, 44, 48, 149, 153, 157, 161, 165 };
        //int[] Channel_5G_40M = { 36, 40, 44, 48, 149, 153, 157, 161 };
        //int[] Channel_5G_80M = { 36, 40, 44, 48, 149, 153, 157, 161 };
        //string[] Mode_2_4G = { "11N_20M", "11N_40M", "HT40+", "HT40-"};
        //string[] Mode_5G = { "11N_20M", "11N_40M", "11AC_20M", "11AC_40M", "11AC_80M", "HT40+", "HT40-"};

        //struct WirelessParameter
        //{
        //    public string band;
        //    public string mode;
        //    public string ssid_config;
        //    public string ssid_text;
        //    public string channel_config;
        //    public int channel_start;
        //    public int channel_stop;
        //    public string security_config;
        //    public string security_mode;
        //    public string key_index;
        //    public string passphrase;            
        //}

        //int PingTimeout = 1000;

        //string model;
        //string security = "None";
        //string passphrase = string.Empty;

        //Thread threadRouterFT;
        //bool bRouterFTThreadRunning = false;

        //public string DutConfigurationType = string.Empty;
        //public string DutConfigurationFile = string.Empty;        

        ///* Excel object */
        //Excel.Application excelAppRouter;
        //Excel.Workbook excelWorkBookRouter;
        //Excel.Worksheet excelWorkSheetRouter;
        //Excel.Range excelRangeRouter;

        //object misValue = System.Reflection.Missing.Value;

        //public class WifiBasic
        //{
        //    public string band;
        //    public string mode;
        //    public int channel;
        //    public string ssid;
        //    public string security;
        //    public string passphrase;
        //}

        //public class AtteuationValue
        //{
        //    public int start;
        //    public int stop;
        //    public int steps;
        //}

        //public class TstFile
        //{
        //    public string txTst;
        //    public string rxTst;
        //    public string biTst;
        //}


        #region global variable
        int i_ExcelRowsLimit = 10000;
        #endregion 
        ///*============================================================================================*/
        ///*================================= Excel Function Area   ====================================*/
        ///*============================================================================================*/
        #region Excel Function Area

        private void releaseExcelObject(object obj)
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

        #endregion Excel Function Area


        //private void saveExcelRouter(string savePath)
        //{
        //    try
        //    {
        //        excelWorkBookRouter.SaveAs(savePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine(ex);
        //        MessageBox.Show("Error: Total of file path is exceeds 218 characters.");
        //    }
        //}

        //private void Save_and_Close_Excel_File()
        //{
        //    try
        //    {
        //        if (excelWorkBookCommon != null)
        //        {
        //            /* Save excel data */
        //            excelWorkBookCommon.Save();
        //            closeExcelCommon();
        //        }
        //    }
        //    catch { }

        //}

        //private void closeExcelCommon()
        //{
        //    /* Turn on interactive mode */
        //    excelAppCommon.Interactive = true;
        //    excelWorkBookCommon.Close();
        //    excelAppCommon.Quit();

        //    /*System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppRfi);*/
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkSheetCommon);
        //    excelWorkSheetCommon = null;
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkBookCommon);
        //    excelWorkBookCommon = null;
        //    releaseObject_Common(excelRangeCommon);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelRangeCommon);
        //    excelRangeCommon = null;
        //    releaseObject_Common(excelAppCommon);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppCommon);
        //    excelAppCommon = null;
        //    GC.Collect();
        //}


        //private void SaveEndTime_and_IEversion_to_Excel(CBT_SeleniumApi csUSBStorageBrowser)
        //{
        //    try
        //    {
        //        excelWorkSheetCommon = excelWorkBookCommon.Sheets[1];

        //        excelAppCommon.Cells[10, 4] = "Script Path";
        //        excelAppCommon.Cells[10, 5] = txtUSBStorageFunctionTestScriptFilePath.Text;
        //        excelRangeCommon = excelWorkSheetCommon.Cells[10, 5];
        //        SettingExcelAlignment(excelRangeCommon, "H", "R");
        //        excelAppCommon.Cells[11, 4] = "Browser Version";


        //        excelAppCommon.Cells[11, 5] = csUSBStorageBrowser.BrowserVersion();

        //        excelRangeCommon = excelWorkSheetCommon.Cells[11, 5];
        //        SettingExcelAlignment(excelRangeCommon, "H", "R");



        //        excelAppCommon.Cells[7, 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");  // Save End Time to Excel
        //    }
        //    catch { }

        //}       

        //private void closeExcelRouter()
        //{
        //    /* Turn on interactive mode */
        //    excelAppRouter.Interactive = true;
        //    excelWorkBookRouter.Close();
        //    excelAppRouter.Quit();

        //    /*System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppRfi);*/
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkSheetRouter);
        //    excelWorkSheetRouter = null;
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkBookRouter);
        //    excelWorkBookRouter = null;
        //    releaseObject_Router(excelRangeRouter);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelRangeRouter);
        //    excelRangeRouter = null;
        //    releaseObject_Router(excelAppRouter);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppRouter);
        //    excelAppRouter = null;
        //    GC.Collect();
        //}

        //private void releaseObject_Router(object obj)
        //{
        //    try
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        //        obj = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        obj = null;
        //        MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
        //    }
        //    finally
        //    {
        //        GC.Collect();
        //    }
        //}

        //private void SettingExcelAlignment(Excel.Range excelRangeCGA2121, string direction, string AlignmentType)
        //{
        //    //--- [direction] H:水平方向 V:垂直方向 (可同時設定多個屬性)   ex:"HV"、"H"、"V"
        //    //--- [AlignmentType] C:置中 L:靠左 R:靠右

        //    if (AlignmentType.CompareTo("C") == 0)      // 置中
        //    {
        //        if (direction.ToLower().IndexOf("h") >= 0)
        //            excelRangeCGA2121.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        //        if (direction.ToLower().IndexOf("v") >= 0)
        //            excelRangeCGA2121.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
        //    }
        //    else if (AlignmentType.CompareTo("L") == 0) // 靠左
        //    {
        //        if (direction.ToLower().IndexOf("h") >= 0)
        //            excelRangeCGA2121.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
        //        if (direction.ToLower().IndexOf("v") >= 0)
        //            excelRangeCGA2121.VerticalAlignment = Excel.XlHAlign.xlHAlignLeft;
        //    }
        //    else if (AlignmentType.CompareTo("R") == 0) // 靠右
        //    {
        //        if (direction.ToLower().IndexOf("h") >= 0)
        //            excelRangeCGA2121.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
        //        if (direction.ToLower().IndexOf("v") >= 0)
        //            excelRangeCGA2121.VerticalAlignment = Excel.XlHAlign.xlHAlignRight;
        //    }
        //}


        //#endregion

        
        
        
        //private string createRouterSubFolder(string ModelName)
        //{
        //    return ((ModelName == "") ? "Router_" : ModelName + "_") + DateTime.Now.ToString("yyyyMMddHHmmss");
        //}

        ///* delegate call back function */
        //private void SetText(string msg, TextBox txtInformation)
        //{
        //    txtInformation.AppendText(msg + Environment.NewLine);
        //}

        //private void SetTextC(string msg, TextBox txtInformation)
        //{
        //    txtInformation.AppendText(msg);
        //}

        //private bool CsvAppend(string filePath, string str)
        //{
        //    if (!File.Exists(filePath))
        //    {
        //        Debug.WriteLine("File doesn't exist");
        //        return false;
        //    }

        //    try
        //    {
        //        StringBuilder csv = new StringBuilder();
        //        csv.AppendLine(str);
        //        File.AppendAllText(filePath, csv.ToString());
        //    }
        //    catch (Exception ex)
        //    {

        //        Debug.WriteLine("Write CSV Failed");
        //        return false;   
        //    }

        //    return true;
        //}

        //private bool InserLine2TextFile(string filePath, int iInsertLine, string text)
        //{
        //    if (!File.Exists(filePath)) return false;
        //    //string sTestFileName = @"e:\t1.txt";
        //    //int iInsertLine = 5;
        //    //string text = "插入的内容";
        //    string sText = "";
        //    System.IO.StreamReader sr = new System.IO.StreamReader(filePath);

        //    int iLnTmp = 0; //记录文件行数
        //    while (!sr.EndOfStream)
        //    {
        //        iLnTmp++;
        //        if (iLnTmp == iInsertLine)
        //        {
        //            sText += text + "\r\n";  //将值插入
        //        }
        //        string sTmp = sr.ReadLine();    //记录当前行
        //        sText += sTmp + "\r\n";
        //    }
        //    sr.Close();

        //    System.IO.StreamWriter sw = new System.IO.StreamWriter(filePath, false);
        //    sw.Write(sText);
        //    sw.Flush();
        //    sw.Close();
        //    return true;
        //}

        //private string ThroughputValue(string filePath)
        //{
        //    if (!File.Exists(filePath)) return null;
                        
        //    string sText = "";
        //    System.IO.StreamReader sr = new System.IO.StreamReader(filePath);

        //    while (!sr.EndOfStream)
        //    {
        //        string sTmp = sr.ReadLine();    //记录当前行
        //        sText = sTmp;
        //    }
        //    sr.Close();

        //    string[] t = sText.Split(',');
        //    return t[9];
        //}

        //private string ThroughputUnit(string filePath)
        //{
        //    if (!File.Exists(filePath)) return null;

        //    string sText = "";
        //    string preText = "";
        //    System.IO.StreamReader sr = new System.IO.StreamReader(filePath);

        //    while (!sr.EndOfStream)
        //    {
        //        string sTmp = sr.ReadLine();    //记录当前行
        //        preText = sText;
        //        sText = sTmp;
        //    }
        //    sr.Close();

        //    string[] t = preText.Split(',');
        //    return t[9];
        //}

        //private bool ConfigureRouter(string modelName, string hostIP, string userName, string passWord, WifiBasic wifi)
        //{
        //    switch (modelName.ToLower())
        //    {
        //        case "tch-2.4g":
        //            return Tch24GDeviceConfigure(hostIP, userName, passWord, wifi);                    
        //        case "tch-dual":
        //            return Tch5GDeviceConfigure(hostIP, userName, passWord, wifi);
        //        default:
        //            return false;                    
        //    }            
        //}

        //private bool WebHttpGetData(string host, string agent, bool keepAlive, bool Auth, string username, string password, string compareText, ref string response, string contentType = "")
        //{
        //    string _uri = host;
        //    string _strHtml = string.Empty;
            
        //    try
        //    {
        //        WebRequest myRequest = WebRequest.Create(_uri);
        //        myRequest.Timeout = 30000;
        //        if (Auth)
        //            myRequest.Credentials = new NetworkCredential(username, password);

        //        HttpWebRequest myHttpWebRequest = (HttpWebRequest)myRequest;
        //        myHttpWebRequest.Timeout = 30000;
        //        myHttpWebRequest.Method = "GET";
        //        myHttpWebRequest.UserAgent = agent;
        //        myHttpWebRequest.KeepAlive = keepAlive;

        //        if (contentType != "")
        //            myHttpWebRequest.ContentType = contentType;

        //        using (WebResponse myWebResponse = myHttpWebRequest.GetResponse())
        //        {
        //            using (Stream myStream = myWebResponse.GetResponseStream())
        //            {
        //                using (StreamReader myReader = new StreamReader(myStream))
        //                {
        //                    _strHtml = myReader.ReadToEnd();
        //                    //document.Load(myReader.ReadToEnd());
        //                }
        //            }
        //            myWebResponse.Close();
        //        }

        //        //textBox1.Text = _strHtml;
        //        response = _strHtml;
        //        //response = document.DocumentNode.InnerText;

        //        return true;
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.ToString());
        //        return false;
        //    }
        //}

        //private bool WebHttpPostData(string host, string data, string agent, string ContentType, bool keepAlive, string Type, bool Auth, string username, string password, ref string response)
        //{
        //    byte[] bs = Encoding.ASCII.GetBytes(data);

        //    HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(host);
        //    req.Method = "POST";
        //    req.ContentType = ContentType;
        //    req.KeepAlive = keepAlive;
        //    //req.ProtocolVersion = HttpVersion.Version11;
        //    req.UserAgent = agent;

        //    req.ContentLength = bs.Length;
        //    if (Type == "jnap")
        //    {
        //        //req.Headers.Add("
        //        //req.Headers.Add("x-jnap-authorization", "Basic YWRtaW46YWRtaW4=\r\n");
        //        //req.Headers.Add("x-jnap-action", "http://linksys.com/jnap/core/Transaction\r\n");
        //        //req.Headers.Add("x-requested-with", "XMLHttpRequest\r\n");
        //    }

        //    if (Auth)
        //    {
        //        req.Headers.Add("Authorization", "Basic YWRtaW46YWRtaW4=\r\n");
        //        req.Headers.Add("Credentials", username + ":" + password);
        //        //req.Headers.Add("Credentials", "admin:admin");
        //    }

        //    try
        //    {
        //        using (Stream reqStream = req.GetRequestStream())
        //        {
        //            reqStream.Write(bs, 0, bs.Length);
        //        }
        //        using (WebResponse wr = req.GetResponse())
        //        {
        //            Stream myStream = wr.GetResponseStream();
        //            StreamReader reader = new StreamReader(myStream);
        //            string strHtml = reader.ReadToEnd();
        //            //reader
        //            //textBox1.AppendText(strHtml);
        //            response = strHtml;

        //            //textBox1.AppendText(Environment.NewLine);
        //            //textBox1.AppendText(Environment.NewLine);


        //            if (strHtml.ToLower().IndexOf("error") >= 0)
        //            {
        //                //textBox1.AppendText("Perfect");
        //                return false;
        //            }
        //            wr.Close();
        //            return true;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //    }
        //}
        
        ///*============================================================================================*/
        ///*================================  Delegate function ========================================*/
        ///*============================================================================================*/
        //private void showWirelessPatameterAllContent(WirelessParameter wp, TextBox textbox)
        //{
        //    string str = string.Empty;
        //    str += "band: " + wp.band;
        //    str += "\tmode: " + wp.mode;
        //    str += "\tssid_config: " + wp.ssid_config;
        //    str+= "\tssid_text: " + wp.ssid_text ;
        //    str += "\tchannel_config: " + wp.channel_config;
        //    str += "\tchannel_start: " + wp.channel_start;
        //    str += "\tchannel_stop: " + wp.channel_stop;
        //    str += "\tsecurity_config: " + wp.security_config;
        //    str += "\tsecurity_mode: " + wp.security_mode;
        //    str += "\tkey_index: " + wp.key_index;
        //    str += "\tpassphrase: " + wp.passphrase;         

        //    textbox.AppendText(str + Environment.NewLine);
        //}

        //private void showWirelessPatameterPartContent(WirelessParameter wp, TextBox textbox)
        //{
        //    string str = string.Empty;
        //    str += "band: " + wp.band;
        //    str += "\tmode: " + wp.mode;
        //    str += "\tssid_config: " + wp.ssid_config;
        //    str += "\tssid_text: " + wp.ssid_text;
        //    str += "\tchannel_config: " + wp.channel_config;
        //    str += "\tchannel_start: " + wp.channel_start;
        //    str += "\tchannel_stop: " + wp.channel_stop;
        //    str += "\tsecurity_config: " + wp.security_config;
        //    str += "\tsecurity_mode: " + wp.security_mode;
        //    str += "\tkey_index: " + wp.key_index;
        //    str += "\tpassphrase: " + wp.passphrase;

        //    textbox.AppendText(str + Environment.NewLine);
        //}

        ///*============================================================================================*/
        ///*============================= Print Screen Function Area   =================================*/
        ///*============================================================================================*/
        //#region Print Screen Function Area

        //private void PrintFullScreen_Common(string FilePath)
        //{
        //    Bitmap myImage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
        //    Graphics g = Graphics.FromImage(myImage);
        //    g.CopyFromScreen(new Point(0, 0), new Point(0, 0), new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
        //    IntPtr dc1 = g.GetHdc();
        //    g.ReleaseHdc(dc1);
        //    myImage.Save(FilePath);
        //}

        //private void PrintFocusedWindow_Common(string FilePath)
        //{
        //    Bitmap myImage = new Bitmap(this.Width, this.Height);
        //    Graphics g = Graphics.FromImage(myImage);
        //    g.CopyFromScreen(new Point(this.Location.X, this.Location.Y), new Point(0, 0), new Size(this.Width, this.Height));
        //    IntPtr dc1 = g.GetHdc();
        //    g.ReleaseHdc(dc1);
        //    myImage.Save(FilePath);
        //}
        
        //#endregion

        ///*============================================================================================*/
        ///*=============================== GUI Test Function Area   ===================================*/
        ///*============================================================================================*/

        //#region

        ///* Declare delegate prototype */
        ////public delegate void showHddtaRfiRestartDelegate(string str);
        ////public delegate void showHddtaRfiChannelPowerDelegate(string ch_power, int frequency);
        ////public delegate void showHddtaRfiResultDelegate(int CenterFrequency, double Level, string Constellation, string ReportPwr, string Snr, string CE, string UE);

        ///* Declare delegate prototype */
        ////private delegate void SetTextCallBack(string text, TextBox textbox);
        

        //public struct ModelGroupStruct
        //{
        //    public string ModelName;
        //    public string SN;
        //    public string SWver;
        //    public string HWver;

        //    public ModelGroupStruct(string ModelName, string SN, string SWver, string HWver)
        //    {
        //        this.ModelName = ModelName;
        //        this.SN = SN;
        //        this.SWver = SWver;
        //        this.HWver = HWver;
        //    }
        //}

        //struct TestBrowerSetting
        //{
        //    public string TestRun;
        //    public bool StopWhenTestError;
        //}

        //struct LoginSetting
        //{
        //    public string GatewayIP;
        //    public string UserName;
        //    public string Password;
        //    public string HTTP_Port;
        //    public int RebootWaitTime;
        //    public bool ConsoleLog;
        //}

        //struct struct_ScriptDataUSBStorage
        //{
        //    public string TestIndex;
        //    public string Action;
        //    public string ActionName;
        //    public string ElementType;
        //    //public string WriteExpectedValue;
        //    //public string WriteExpectedValue;
        //    public string RadioButtonExpectedValueXpath;
        //    public string WriteExpectedValue;
        //    public string ElementXpath;
        //    public string TestTimeOut;
        //    public string GetValue;
        //}

        ////struct ScriptData
        ////{
        ////    public string TestIndex;
        ////    public string TestStep;
        ////    public string Action;
        ////    public string ActionName;
        ////    public string ElementType;
        ////    public string WriteValue;
        ////    public string ExpectedValue;
        ////    public string RadioButtonExpectedValueXpath;
        ////    public string URL;
        ////    public string ElementXpath;
        ////    public string ApplyButtonXpath;
        ////    public string TestTimeOut;
        ////    public string TriggerReboot;
        ////    public string GetValue;
        ////}

        //private void closeExcelReadFileCommon()
        //{
        //    /* Turn on interactive mode */
        //    excelReadAppCommon.Interactive = true;
        //    excelReadWorkBookCommon.Close();
        //    excelReadAppCommon.Quit();

        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelReadWorkSheetCommon);
        //    excelReadWorkSheetCommon = null;
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelReadWorkBookCommon);
        //    excelReadWorkBookCommon = null;
        //    releaseObject_Common(excelReadRangeCommon);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelReadRangeCommon);
        //    excelReadRangeCommon = null;
        //    releaseObject_Common(excelReadAppCommon);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelReadAppCommon);
        //    excelReadAppCommon = null;
        //    GC.Collect();
        //}

        //private void releaseObject_Common(object obj)
        //{
        //    try
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        //        obj = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        obj = null;
        //        MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
        //    }
        //    finally
        //    {
        //        GC.Collect();
        //    }
        //}


        //#endregion

        ///*============================================================================================*/
        ///*============================== USB Storage Function Area ===================================*/
        ///*============================================================================================*/

        //#region USB Storage

        //public delegate void showCommonGUIDelegate();

        //Thread threadCommonFT;
        //Thread threadCommonFTstopEvent;
        //Thread threadCommonFTscriptItem;
        //bool bCommonFTThreadRunning = false;
        //bool bCommonTestComplete = true;
        //bool bCommonSingleScriptItemRunning = false;

        public struct ModelInfo
        {
            public string ModelName;
            public string SN;
            public string SwVersion;
            public string HwVersion;
        }

        ModelInfo modelInfo;

        public delegate void ShowTextboxContentDelegate(string str, TextBox textbox);


        ///* Excel object */
        //Excel.Application excelAppCommon = null;
        //Excel.Workbook excelWorkBookCommon = null;
        //Excel.Worksheet excelWorkSheetCommon = null;
        //Excel.Range excelRangeCommon = null;

        //Excel.Application excelReadAppCommon = null;
        //Excel.Workbook excelReadWorkBookCommon = null;
        //Excel.Worksheet excelReadWorkSheetCommon = null;
        //Excel.Range excelReadRangeCommon = null;

        //private string ReadConsoleLogByBytesUSBStorage(Comport com)
        //{
        //    int byteToRead = com.GetBytesToRead();
        //    return com.Read(0, byteToRead);
        //}

        //private void saveExcelUSBStorage(string savePath)
        //{
        //    try
        //    {
        //        excelWorkBookCommon.SaveAs(savePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine(ex);
        //        MessageBox.Show("Error: Check the file path must within 218 characters.");
        //    }
        //}


        //#endregion

        ///*============================================================================================*/
        ///*=================================== New Added Function =====================================*/
        ///*============================================================================================*/

        //private void MessageBoxTopMost(string Title, string Info) // 將 MessageBox訊息置頂
        //{
        //    MessageBox.Show(new Form { TopMost = true }, Info, Title, MessageBoxButtons.OK);
        //}

        //private void HideAllTabPage()
        //{
        //    this.tpWebGUI_TestCondition.Parent = null;      // hide tpWebGUI_TestCondition TabPage
        //    this.tpWebGUI_FunctionTest.Parent = null;       // hide tpWebGUI_FunctionTest TabPage
        //    this.tpUSBStorageFunctionTest.Parent = null;       // hide tpUSBStorage_FunctionTest TabPage
        //    this.SallyTestPage.Parent = null;               //hide SallyTestPage TabPage
        //}

        private string LoadFile_Common(string sFileName, string sFilter, string InitialDirectory)
        {
            try
            {
                string sFilesName = string.Empty;
                //string filename = @"MibReadWriteTestCondition.xlsx";            
                // Displays a SaveFileDialog so the user can select the exported test condition excel file
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                //openFileDialog1.Multiselect = true;
                openFileDialog1.FileName = sFileName;
                //openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\exportData\";
                openFileDialog1.InitialDirectory = InitialDirectory;
                // Set filter for file extension and default file extension
                //openFileDialog1.Filter = "XLSX file|*.xlsx";
                openFileDialog1.Filter = sFilter;

                // If the file name is not an empty string open it for opening.
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName != "")
                {
                    sFilesName = openFileDialog1.FileName;
                }

                return sFilesName;
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        private void SaveLog(TextBox textbox)
        {
            SaveFileDialog saveDebugFileDialog = new SaveFileDialog();

            saveDebugFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            saveDebugFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveDebugFileDialog.FilterIndex = 1;
            saveDebugFileDialog.RestoreDirectory = true;
            saveDebugFileDialog.FileName = "Log" + DateTime.Now.ToString("yyyyMMdd-HHmmss");

            if (saveDebugFileDialog.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText(saveDebugFileDialog.FileName, textbox.Text);
            }
        }

        private void SaveLog(string savePath, TextBox textbox)
        {
            SaveFileDialog saveDebugFileDialog = new SaveFileDialog();

            saveDebugFileDialog.InitialDirectory = savePath;
            saveDebugFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveDebugFileDialog.FilterIndex = 1;
            saveDebugFileDialog.RestoreDirectory = true;
            saveDebugFileDialog.FileName = "Log" + DateTime.Now.ToString("yyyyMMdd-HHmmss");

            if (saveDebugFileDialog.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText(saveDebugFileDialog.FileName, textbox.Text);
            }
        }

        

        private bool ConvertExcelToDualArray(string sExcelFile, ref string[,] saCondition)
        {
            if (sExcelFile == null || sExcelFile == "" || !File.Exists(sExcelFile))
            {
                throw new ArgumentNullException();
            }

            int iRealCount = 0;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(sExcelFile);
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            if (rowCount >= i_ExcelRowsLimit) rowCount = i_ExcelRowsLimit;

            string[,] rowdata = new string[rowCount, colCount]; //Added  for compare index 

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                bool bHasData = false;
                for (int j = 1; j <= colCount; j++)
                {
                    ////new line
                    //if (j == 1)
                    //    Console.Write("\r\n");

                    //string s1 = xlRange.Cells[i, j]
                    //string s2 = xlRange.Cells[i, j].Value2;

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        rowdata[i - 1, j - 1] = xlRange.Cells[i, j].Value2.ToString();
                        bHasData = true;
                    }
                    else
                    {
                        //rowdata[i - 1, j - 1] = "";
                    }
                    //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                }

                if (bHasData)
                    iRealCount++;

                if (rowdata[i - 1, 0] != null && rowdata[i - 1, 0].ToLower() == "end")
                    break;
            }

            saCondition = rowdata;            

            /* Close Excel */
            /* Turn on interactive mode */
            xlApp.Interactive = true;
            //xlWorkbook.Close();  //Mark by Jin on 2019/01/31 because chamberperformance test bed will hang when go on this step.
            xlApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
            xlWorksheet = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;
            releaseExcelObject(xlRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            xlRange = null;
            releaseExcelObject(xlApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            GC.Collect();

            return true;
        }

        private bool ConvertDatagridViewDataToDualArray(DataGridView dgvData, ref string[,] saCondition)
        {
            string[,] saData = new string[dgvData.RowCount - 1, dgvData.ColumnCount];

            for (int iRowIndex = 0; iRowIndex < dgvData.RowCount - 1; iRowIndex++)
            {
                for (int iColumnIndex = 0; iColumnIndex < dgvData.ColumnCount; iColumnIndex++)
                {
                    saData[iRowIndex, iColumnIndex] = dgvData.Rows[iRowIndex].Cells[iColumnIndex].Value.ToString();
                }
            }

            saCondition = saData;

            return true;
        }

        public void ShowTextboxContent(string str, TextBox textbox)
        {
            textbox.Text = str;
        }
        

        /*============================================================================================*/
        /*=============================== Comport Loading Function  Area =============================*/
        /*============================================================================================*/
        #region
        public bool ReadComportSettingAndInitial()
        {
            string FileName = System.Windows.Forms.Application.StartupPath + "\\config\\SerialPort.xml";
            Comport comport = null;

            XmlDocument doc = new XmlDocument();
            doc.Load(FileName);

            XmlNode nodeXml = doc.SelectSingleNode("SerialPortSetting");
            if (nodeXml == null)
                return false;
            XmlElement element = (XmlElement)nodeXml;
            string strID = element.GetAttribute("Item");
            Debug.WriteLine(strID);
            if (strID.CompareTo("Serial Port Setting") != 0)
            {
                MessageBox.Show("This XML file is incorrect.", "Error");
                return false;
            }

            ///
            /// Read Function Test configuration settings
            ///

            XmlNode nodeComportSetting = doc.SelectSingleNode("/SerialPortSetting/SerialPortParameter");

            try
            {
                string PortNum = nodeComportSetting.SelectSingleNode("PortNum").InnerText;
                string BaudRate = nodeComportSetting.SelectSingleNode("BaudRate").InnerText;
                string Parity = nodeComportSetting.SelectSingleNode("Parity").InnerText;
                string DataBit = nodeComportSetting.SelectSingleNode("DataBit").InnerText;
                string StopBit = nodeComportSetting.SelectSingleNode("StopBit").InnerText;
                string FlowControl = nodeComportSetting.SelectSingleNode("FlowControl").InnerText;
                string ReadTimeout = nodeComportSetting.SelectSingleNode("ReadTimeout").InnerText;
                string WriteTimeout = nodeComportSetting.SelectSingleNode("WriteTimeout").InnerText;


                Debug.WriteLine("PortNum: " + PortNum);
                Debug.WriteLine("BaudRate: " + BaudRate);
                Debug.WriteLine("Parity: " + Parity);
                Debug.WriteLine("DataBit: " + DataBit);
                Debug.WriteLine("StopBit: " + StopBit);
                Debug.WriteLine("FlowControl: " + FlowControl);
                Debug.WriteLine("ReadTimeout: " + ReadTimeout);
                Debug.WriteLine("WriteTimeout: " + WriteTimeout);

                comport = new Comport();

                if (comport.isOpen() == true)
                {
                    comport.Close();
                }

                //comport.init(cbPort.Text, cbBaudrate.Text, cbParity.Text, cbData.Text, cbStop.Text, cbFlow.Text, tbReadTimeOut.Text, tbWriteTimeOut.Text);
                bool result = comport.init(PortNum, Convert.ToInt32(BaudRate), Parity, Convert.ToInt32(DataBit),
                    StopBit, FlowControl, Convert.ToInt32(ReadTimeout), Convert.ToInt32(WriteTimeout));
                //if(result == true) 

                comport.Open();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/SerialPortSetting/SerialPortParameter " + ex);
            }

            return true;
        }

        public bool ReadComport2SettingAndInitial()
        {
            string FileName = System.Windows.Forms.Application.StartupPath + "\\config\\SerialPort2.xml";
            Comport2 comport2 = null;

            XmlDocument doc = new XmlDocument();
            doc.Load(FileName);

            XmlNode nodeXml = doc.SelectSingleNode("SerialPort2Setting");
            if (nodeXml == null)
                return false;
            XmlElement element = (XmlElement)nodeXml;
            string strID = element.GetAttribute("Item");
            Debug.WriteLine(strID);
            if (strID.CompareTo("Serial Port2 Setting") != 0)
            {
                MessageBox.Show("This XML file is incorrect.", "Error");
                return false;
            }

            ///
            /// Read Function Test configuration settings
            ///

            XmlNode nodeComportSetting = doc.SelectSingleNode("/SerialPort2Setting/SerialPortParameter");

            try
            {
                string PortNum = nodeComportSetting.SelectSingleNode("PortNum").InnerText;
                string BaudRate = nodeComportSetting.SelectSingleNode("BaudRate").InnerText;
                string Parity = nodeComportSetting.SelectSingleNode("Parity").InnerText;
                string DataBit = nodeComportSetting.SelectSingleNode("DataBit").InnerText;
                string StopBit = nodeComportSetting.SelectSingleNode("StopBit").InnerText;
                string FlowControl = nodeComportSetting.SelectSingleNode("FlowControl").InnerText;
                string ReadTimeout = nodeComportSetting.SelectSingleNode("ReadTimeout").InnerText;
                string WriteTimeout = nodeComportSetting.SelectSingleNode("WriteTimeout").InnerText;


                Debug.WriteLine("PortNum: " + PortNum);
                Debug.WriteLine("BaudRate: " + BaudRate);
                Debug.WriteLine("Parity: " + Parity);
                Debug.WriteLine("DataBit: " + DataBit);
                Debug.WriteLine("StopBit: " + StopBit);
                Debug.WriteLine("FlowControl: " + FlowControl);
                Debug.WriteLine("ReadTimeout: " + ReadTimeout);
                Debug.WriteLine("WriteTimeout: " + WriteTimeout);

                comport2 = new Comport2();

                if (comport2.isOpen() == true)
                {
                    comport2.Close();
                }

                //comport.init(cbPort.Text, cbBaudrate.Text, cbParity.Text, cbData.Text, cbStop.Text, cbFlow.Text, tbReadTimeOut.Text, tbWriteTimeOut.Text);
                bool result = comport2.init(PortNum, Convert.ToInt32(BaudRate), Parity, Convert.ToInt32(DataBit),
                    StopBit, FlowControl, Convert.ToInt32(ReadTimeout), Convert.ToInt32(WriteTimeout));
                //if(result == true) 

                comport2.Open();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/SerialPortSetting/SerialPortParameter " + ex);
            }

            return true;
        }

        #endregion

        /*============================================================================================*/
        /*============================ DataGridView Process Function  Area ===========================*/
        /*============================================================================================*/
        #region

        /* https://dotblogs.com.tw/abbee/archive/2010/09/29/17982.aspx */
        /// <summary>
        /// 將欄位填滿Grid
        /// </summary>
        /// <param name="grid">欲設定的grid</param>
        /// <param name="Average">平均分配欄寬</param>
        public static void FillAllColumns(DataGridView grid, bool Average)
        {
            foreach (DataGridViewColumn clm in grid.Columns)
            {
                if (!Average && clm.ValueType.IsValueType && clm != grid.Columns[grid.Columns.Count - 1])
                { clm.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader; }
                else { clm.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; }
            }
        }

        /// <summary>
        /// 設置對齊方式
        /// </summary>
        /// <param name="align"></param>
        /// <returns></returns>
        public static HorizontalAlignment TransAlign(DataGridViewContentAlignment align)
        {
            switch (align)
            {
                case DataGridViewContentAlignment.TopLeft:
                case DataGridViewContentAlignment.MiddleLeft:
                case DataGridViewContentAlignment.BottomLeft:
                    return HorizontalAlignment.Left;

                case DataGridViewContentAlignment.TopCenter:
                case DataGridViewContentAlignment.MiddleCenter:
                case DataGridViewContentAlignment.BottomCenter:
                    return HorizontalAlignment.Center;

                case DataGridViewContentAlignment.TopRight:
                case DataGridViewContentAlignment.MiddleRight:
                case DataGridViewContentAlignment.BottomRight:
                    return HorizontalAlignment.Right;
            }
            return HorizontalAlignment.Left;
        }

        /// <summary>
        /// 檢查該列是否有未填寫欄位
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static bool IsRowHasNull(DataGridViewRow row)
        {
            foreach (DataGridViewCell cell in row.Cells)
            { if (Convert.ToString(cell.Value) == string.Empty) { return true; } }
            return false;
        }

        /// <summary>
        /// 按傳入欄位依序排序
        /// </summary>
        /// <param name="grid"></param>
        /// <param name="Columns"></param>
        public static void SetColumnSort(DataGridView grid, params string[] Columns)
        {
            List<String> clmns = new List<String>(Columns);
            foreach (DataGridViewColumn clm in grid.Columns)
            {
                if (clmns.Contains(clm.Name)) { clm.DisplayIndex = clmns.IndexOf(clm.Name); }
                else { clm.DisplayIndex = grid.Columns.Count - 1; }
            }
        }

        /* DataGridView 欄位參考
         http://www.cnblogs.com/scottckt/archive/2007/10/09/917874.html
        */

        /// <summary>
        /// 將所選的列上移一列
        /// </summary>
        /// <param name="grid"></param>
        /// <param name="Columns"></param>
        private void DataGridViewMoveUp(DataGridView grid)
        {
            if (grid.RowCount <= 1) return;

            int iDgvSeletedIndex = grid.CurrentRow.Index;

            if (iDgvSeletedIndex == 0 || iDgvSeletedIndex == grid.RowCount - 1) return; //Top or Empty Button

            //DataGridViewRow row = (DataGridViewRow)grid.Rows[iDgvSeletedIndex].Clone();
            DataGridViewRow row = (DataGridViewRow)grid.Rows[0].Clone();

            for (Int32 index = 0; index < row.Cells.Count; index++)
            {
                row.Cells[index].Value = grid.Rows[iDgvSeletedIndex].Cells[index].Value;
            }
            //return clonedRow;


            grid.Rows.Insert(iDgvSeletedIndex - 1, row);
            grid.Rows.RemoveAt(iDgvSeletedIndex + 1);
            //grid.CurrentCell = grid[0, iDgvSeletedIndex - 1];
            DataGridViewCell cellNow = grid.CurrentCell;

            grid.CurrentCell = grid[cellNow.ColumnIndex, iDgvSeletedIndex - 1];
        }

        /// <summary>
        /// 將所選的列下移一列
        /// </summary>
        /// <param name="grid"></param>
        /// <param name="Columns"></param>
        private void DataGridViewMoveDown(DataGridView grid)
        {
            if (grid.RowCount <= 1) return;

            int iDgvSeletedIndex = grid.CurrentRow.Index;

            if (iDgvSeletedIndex >= grid.RowCount - 2) return; //Bottom

            DataGridViewRow row = (DataGridViewRow)grid.Rows[0].Clone();
            for (Int32 index = 0; index < row.Cells.Count; index++)
            {
                row.Cells[index].Value = grid.Rows[iDgvSeletedIndex + 1].Cells[index].Value;
            }

            grid.Rows.Insert(iDgvSeletedIndex, row);
            grid.Rows.RemoveAt(iDgvSeletedIndex + 2);
            //grid.CurrentCell = grid[0, iDgvSeletedIndex + 1];
            DataGridViewCell cellNow = grid.CurrentCell;

            grid.CurrentCell = grid[cellNow.ColumnIndex, iDgvSeletedIndex + 1];

            //grid.Rows.

            //grid.Rows.Insert(iDgvSeletedIndex - 1, row);
            //grid.Rows.RemoveAt(iDgvSeletedIndex + 1);
            //grid.CurrentCell = grid[0, iDgvSeletedIndex - 1];
        }

        #endregion

        /*============================================================================================*/
        /*===================================Email Function  AREA ====================================*/
        /*============================================================================================*/

        private void SendReportByGmailWithFile(string sGmailAccount, string sGmailPassword, string[] ReceiverAccount, string sMailSubject, string sMailContent, string AttachedFile)
        {
            try
            {
                //處理有ssl憑證問題的網站出現遠端憑證是無效的錯誤問題
                // 設定 HTTPS 連線時，不要理會憑證的有效性問題
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);

                System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
                //msg.To.Add("blinda12@ms4.hinet.net");
                for (int i = 0; i < ReceiverAccount.Length; i++)
                {
                    msg.To.Add(ReceiverAccount[i]);
                }

                //msg.To.Add("jin.wang@cybertan.com.tw");
                //msg.To.Add("b@b.com");可以發送給多人
                //msg.CC.Add("c@c.com");
                //msg.CC.Add("c@c.com");可以抄送副本給多人 
                //這裡可以隨便填，不是很重要
                msg.From = new MailAddress("XXX@gmail.com", "ATE", System.Text.Encoding.UTF8);
                //msg.From = new MailAddress("XXX@gmail.com", "小魚", System.Text.Encoding.UTF8);
                // msg.From = new MailAddress("yomiyoa1@gmail.com", "小魚", System.Text.Encoding.UTF8);
                /* 上面3個參數分別是發件人地址（可以隨便寫），發件人姓名，編碼*/
                //msg.Subject = "測試標題";//郵件標題
                //msg.Subject = "Test Report";//郵件標題
                msg.Subject = sMailSubject;//郵件標題
                msg.SubjectEncoding = System.Text.Encoding.UTF8;//郵件標題編碼
                //msg.Body = "測試一下"; //郵件內容
                msg.Body = sMailContent; //郵件內容
                msg.BodyEncoding = System.Text.Encoding.UTF8;//郵件內容編碼 
                //msg.Attachments.Add(new Attachment(@"C:\3.png"));  //附件
                msg.Attachments.Add(new Attachment(AttachedFile));  //附件
                msg.IsBodyHtml = true;//是否是HTML郵件 
                //msg.Priority = MailPriority.High;//郵件優先級 

                SmtpClient client = new SmtpClient();
                //client.Credentials = new System.Net.NetworkCredential("XXX@gmail.com", "****"); //這裡要填正確的帳號跟密碼
                //client.Credentials = new System.Net.NetworkCredential("yomiyoa1@gmail.com", "yomiyo22342234"); //這裡要填正確的帳號跟密碼
                client.Credentials = new System.Net.NetworkCredential(sGmailAccount, sGmailPassword); //這裡要填正確的帳號跟密碼
                client.Host = "smtp.gmail.com"; //設定smtp Server
                client.Port = 25; //設定Port
                client.EnableSsl = true; //gmail預設開啟驗證
                client.Send(msg); //寄出信件
                client.Dispose();
                msg.Dispose();
                //MessageBox.Show(this, "郵件寄送成功！");
                //MessageBox.Show(this, "郵件寄送成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SendReportByGmailWithoutFile(string sGmailAccount, string sGmailPassword, string[] ReceiverAccount, string sMailSubject, string sMailContent)
        {
            try
            {
                //處理有ssl憑證問題的網站出現遠端憑證是無效的錯誤問題
                // 設定 HTTPS 連線時，不要理會憑證的有效性問題
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);

                System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
                //msg.To.Add("blinda12@ms4.hinet.net");
                for (int i = 0; i < ReceiverAccount.Length; i++)
                {
                    msg.To.Add(ReceiverAccount[i]);
                }

                //msg.To.Add("jin.wang@cybertan.com.tw");
                //msg.To.Add("b@b.com");可以發送給多人
                //msg.CC.Add("c@c.com");
                //msg.CC.Add("c@c.com");可以抄送副本給多人 
                //這裡可以隨便填，不是很重要
                msg.From = new MailAddress("XXX@gmail.com", "ATE", System.Text.Encoding.UTF8);
                //msg.From = new MailAddress("XXX@gmail.com", "小魚", System.Text.Encoding.UTF8);
                // msg.From = new MailAddress("yomiyoa1@gmail.com", "小魚", System.Text.Encoding.UTF8);
                /* 上面3個參數分別是發件人地址（可以隨便寫），發件人姓名，編碼*/
                //msg.Subject = "測試標題";//郵件標題
                //msg.Subject = "Test Report";//郵件標題
                msg.Subject = sMailSubject;//郵件標題
                msg.SubjectEncoding = System.Text.Encoding.UTF8;//郵件標題編碼
                //msg.Body = "測試一下"; //郵件內容
                msg.Body = sMailContent; //郵件內容
                msg.BodyEncoding = System.Text.Encoding.UTF8;//郵件內容編碼 
                //msg.Attachments.Add(new Attachment(@"C:\3.png"));  //附件
                //msg.Attachments.Add(new Attachment(AttachedFile));  //附件
                msg.IsBodyHtml = true;//是否是HTML郵件 
                //msg.Priority = MailPriority.High;//郵件優先級 

                SmtpClient client = new SmtpClient();
                //client.Credentials = new System.Net.NetworkCredential("XXX@gmail.com", "****"); //這裡要填正確的帳號跟密碼
                //client.Credentials = new System.Net.NetworkCredential("yomiyoa1@gmail.com", "yomiyo22342234"); //這裡要填正確的帳號跟密碼
                client.Credentials = new System.Net.NetworkCredential(sGmailAccount, sGmailPassword); //這裡要填正確的帳號跟密碼
                client.Host = "smtp.gmail.com"; //設定smtp Server
                client.Port = 25; //設定Port
                client.EnableSsl = true; //gmail預設開啟驗證
                client.Send(msg); //寄出信件
                client.Dispose();
                msg.Dispose();
                //MessageBox.Show(this, "郵件寄送成功！");
                //MessageBox.Show(this, "郵件寄送成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //處理有ssl憑證問題的網站出現遠端憑證是無效的錯誤問題
        // 設定 HTTPS 連線時，不要理會憑證的有效性問題
        public static bool ValidateServerCertificate(Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }


        static string MakeFilenameValid(string FN)
        {
            if (FN == null) throw new ArgumentNullException();
            if (FN.EndsWith(".")) FN = Regex.Replace(FN, @"\.+$", "");
            if (FN.Length == 0) throw new ArgumentException();
            if (FN.Length > 245) throw new PathTooLongException();
            foreach (char c in System.IO.Path.GetInvalidFileNameChars()) FN = FN.Replace(c, ' ');

            FN = FN.Replace("/", " ");
            return FN;
        }
        static string MakeFoldernameValid(string FN)
        {
            if (String.IsNullOrEmpty(FN)) throw new ArgumentNullException();
            if (FN.EndsWith(".")) FN = Regex.Replace(FN, @"\.+$", "");
            if (FN.Length == 0) throw new ArgumentException();
            if (FN.Length > 245) throw new PathTooLongException();
            foreach (char c in System.IO.Path.GetInvalidPathChars()) FN = FN.Replace(c, '_');
            return FN.Replace("/", @"\");
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