//---------------------------------------------------------------------------------------//
//  This code was created by CyberTan Sally Lee.                                         // 
//  File           : WebGuiControlFunction.cs                                            // 
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
        private delegate void WebGuiCtrlCommonDelegate_SetTextCallBack(string text, TextBox textbox);
        public delegate void WebGuiCtrlCommonDelegate();

        Thread threadWebGuiCtrlFT;
        Thread threadWebGuiCtrlFTstopEvent;
        Thread threadWebGuiCtrlFTscriptItem;
        bool bWebGuiCtrlFTThreadRunning = false;
        bool bWebGuiCtrlTestComplete = true;
        bool bWebGuiCtrlSingleScriptItemRunning = false;

        CBT_SeleniumApi cs_DutWebGuiCtrlClass = null;
        CBT_SeleniumApi.BrowserType csbt_TestBrowserType;
        CBT_SeleniumApi.BrowserType csbt_TestBrowserChrome = CBT_SeleniumApi.BrowserType.Chrome;
        CBT_SeleniumApi.BrowserType csbt_TestBrowserIE = CBT_SeleniumApi.BrowserType.IE;
        CBT_SeleniumApi.BrowserType csbt_TestBrowserFireFox = CBT_SeleniumApi.BrowserType.FireFox;

        string s_CurrentURL_WebGuiCtrl = string.Empty;
        
        

        //**********************************************************************************//
        //------------------------- Web GUI Control Function Module ------------------------//
        //**********************************************************************************//

        private bool WebDriverInitial()
        {
            if (cs_DutWebGuiCtrlClass == null)
            {
                cs_DutWebGuiCtrlClass = new CBT_SeleniumApi();

                if (!cs_DutWebGuiCtrlClass.init(csbt_TestBrowserType))
                {
                    //CloseExcelReportFile();
                    //CloseExcelReportFile(st_ExcelObjectFinalRepor_WebGuiFwUpDnGrade);
                    //ClodeWebDriver();
                    return false;
                }
                //cs_DutFwGuiBrowser.SettingTimeout(60);
                //cs_DutFwGuiBrowser.WindowMaximize();
                //Thread.Sleep(1000);
                //cs_DutFwGuiBrowser.WindowMinimize();
                Thread.Sleep(3000);
            }

            return true;
        }

        private bool ClodeWebDriver()
        {
            try
            {
                bool bResult = cs_DutWebGuiCtrlClass.Close_WebDriver();
                cs_DutWebGuiCtrlClass = null;
            }
            catch { return false; }
            Thread.Sleep(2000);
            return true;
        }

        private bool CheckWebGuiXPath(ref CBT_SeleniumApi.GuiScriptParameter ScriptPara)
        {
            bool checkXPathResult = false;
            string strInfo = string.Empty;
            Stopwatch waitingtime = new Stopwatch();

            waitingtime.Reset();
            waitingtime.Start();
            do
            {
                checkXPathResult = cs_DutWebGuiCtrlClass.CheckXPathDisplayed(ScriptPara.ElementXpath);
                Thread.Sleep(100);

            } while (waitingtime.ElapsedMilliseconds < Convert.ToInt32(ScriptPara.TestTimeOut) * 1000 && checkXPathResult == false);
            waitingtime.Stop();
            Thread.Sleep(2000);

            return checkXPathResult;
        }

        private bool WebGuiCtrlAction_Set(ref CBT_SeleniumApi.GuiScriptParameter ScriptPara)
        {
            string strInfo = string.Empty;

            try
            {
                cs_DutWebGuiCtrlClass.SetWebElementValue(ref ScriptPara);
            }
            catch (Exception ex)
            {
                strInfo = "-----> Set Value Error:\n" + ex.ToString();
                ScriptPara.Note = strInfo;
                return false;
            }
            return true;
        }

        private bool WebGuiCtrlAction_Get(ref CBT_SeleniumApi.GuiScriptParameter ScriptPara)
        {
            string strInfo = string.Empty;

            try
            {
                Thread.Sleep(2000);
                cs_DutWebGuiCtrlClass.GetWebElementValue(ref ScriptPara);
            }
            catch (Exception ex)
            {
                strInfo = "-----> Get Value Error:\n" + ex.ToString();
                ScriptPara.Note = strInfo;
                return false;
            }
            return true;
        }

    }


}
