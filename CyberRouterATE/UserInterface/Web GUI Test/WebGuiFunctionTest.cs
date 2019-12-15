
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
        CBT_SeleniumApi cs_DutGuiBrowser = null;

        //**********************************************************************************//
        //-------------------------- Web GUI Function Test Event ---------------------------//
        //**********************************************************************************//
        #region Web GUI Test Event

        //private void fWUpgradeDowngradeTestToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    //TestItem = TestItemConstants.TESTITEM_THROUGHPUT;
        //    ToggleToolStripMenuItem(false);
        //    RouterStartPage_tabControl.Hide();

        //    webGUITestToolStripMenuItem.Checked = true;
        //    fWUpgradeDowngradeTestToolStripMenuItem.Checked = true;

        //    tabControl_GUItest.Show();

        //    tsslMessage.Text = tabControl_GUItest.TabPages[tabControl_GUItest.SelectedIndex].Text + " Control Panel";

        //    string xmlFile = System.Windows.Forms.Application.StartupPath + "\\config\\" + Constants.XML_ROUTER_WebGUI;
        //    Debug.WriteLine(File.Exists(xmlFile) ? "File exist" : "File is not existing");
        //}



        //private void tabControl_GUItest_Selected(object sender, TabControlEventArgs e)
        //{
        //    if (tabControl_GUItest.SelectedIndex >= 0)
        //    {
        //        tsslMessage.Text = tabControl_GUItest.TabPages[tabControl_GUItest.SelectedIndex].Text + " Control Panel";
        //        //string str = tabControl_DOCSIStestSeries.SelectedIndex.ToString();
        //        //MessageBox.Show("tabControl_DOCSIStestSeries.SelectedIndex:" + str);
        //    }
        //}

        //==============================================================//
        //========================= Test Code ==========================//
        //==============================================================//
        private void btnSeleniumInitTest_Click(object sender, EventArgs e)
        {
            Thread threadGuiTest;
            threadGuiTest = new Thread(new ThreadStart(DoSeleniumInitTest));
            threadGuiTest.Start();
        }

        private void DoSeleniumInitTest()
        {
            try
            {
                //CBT_SeleniumApi.GuiScriptParameter GuiTestParameter = new CBT_SeleniumApi.GuiScriptParameter();
                //GuiTestParameter.Action
                cs_DutGuiBrowser = new CBT_SeleniumApi();
                cs_DutGuiBrowser.init(CBT_SeleniumApi.BrowserType.IE);
                cs_DutGuiBrowser.SettingTimeout(60);
                cs_DutGuiBrowser.WindowMaximize();
                Thread.Sleep(5000);
                cs_DutGuiBrowser.Close_WebDriver();
            }
            catch (Exception ex)
            {
                string strMessage = "DoSeleniumInitTest() Error:\n" + ex.ToString();
                return;
            }
            MessageBoxTopMost("ATE Information", "Selenium Init Test is Complete!!"); // 將 MessageBox訊息置頂

        }

        #endregion

        //**********************************************************************************//
        //-------------------------- Web GUI Function Test Module --------------------------//
        //**********************************************************************************//

    }
   



}
