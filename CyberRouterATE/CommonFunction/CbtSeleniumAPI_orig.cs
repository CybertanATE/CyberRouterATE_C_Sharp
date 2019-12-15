using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using Keys = OpenQA.Selenium.Keys;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace NS_CbtSeleniumApi_orig
{
    class CBT_SeleniumApi
    {
        //private static ICapabilities BrowserCapabilities;
        //private static IWebDriver driver;

        ICapabilities BrowserCapabilities;
        IWebDriver driver;
        BrowserType browserType;

        string TestBrowser = string.Empty;
        string BrowserDriverProcessesName = string.Empty;


        public CBT_SeleniumApi()
        {
            //GuiTestParameter = new GuiScriptParameter();
        }

        public struct GuiScriptParameter
        {
            public string Index;
            public string TestStep;
            public string Action;
            public string ActionName;
            public string ElementType;
            public string WriteValue;
            public string ExpectedValue;
            public string RadioBtnExpectedValueXpath;
            public string URL;
            public string ElementXpath;
            public string TestTimeOut;
            public string TriggerReboot;
            public string RebootWaitTime;
            public string GetValue;
            public string TestResult;
            public string Note;
        }

        public enum BrowserType
        {
            Chrome = 0,
            IE = 1,
            FireFox = 2,
        }



        public bool init(BrowserType browserType)
        {
            switch (browserType)
            {
                case BrowserType.Chrome:
                    try
                    {
                        TestBrowser = "chrome";
                        BrowserDriverProcessesName = "chromedriver";
                        driver = new ChromeDriver();
                    }
                    catch (Exception info)
                    {
                        MessageBox.Show("****** Some chrome Error happened ******\n\n" + info.ToString());
                        return false;
                    }
                    break;

                case BrowserType.IE:
                    try
                    {
                        TestBrowser = "iexplore";
                        BrowserDriverProcessesName = "IEDriverServer";
                        driver = new InternetExplorerDriver();
                    }
                    catch (Exception info)
                    {
                        string IE_KEY_WORD1 = "browser zoom level";
                        string IE_KEY_WORD2 = "protected mode";
                        if (info.ToString().ToLower().IndexOf(IE_KEY_WORD1) >= 0)
                        {
                            //DesiredCapabilities caps = DesiredCapabilities.InternetExplorer();
                            //caps.SetCapability("EnableNativeEvents", false);
                            //caps.SetCapability("ignoreZoomSetting", true);
                            //driver = new InternetExplorerDriver(caps);
                            MessageBox.Show("Please set IE zoom level to 100% ([Alt+V]--> Zoom--> 100%)");
                        }
                        else if (info.ToString().ToLower().IndexOf(IE_KEY_WORD2) >= 0)
                        {
                            MessageBox.Show("Please enable the IE protected modeset to the same value (enable or disable)\n([Alt+t]--> [O]--> Security)");
                        }
                        else
                        {
                            MessageBox.Show("****** Some IE Error happened ******\n\n" + info);
                        }
                        return false;
                    }
                    break;

                case BrowserType.FireFox:
                    try
                    {
                        TestBrowser = "firefox";
                        BrowserDriverProcessesName = "geckodriver";
                        driver = new FirefoxDriver();
                    }
                    catch (Exception info)
                    {
                        MessageBox.Show("****** Some chrome FireFox happened ******\n\n" + info.ToString());
                        return false;
                    }
                    break;
            }
            BrowserCapabilities = ((RemoteWebDriver)driver).Capabilities;
            return true;
        }

        public bool SetWebElementValue(ref GuiScriptParameter GuiTestParameter)
        {
            IWebElement query;

            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30)); //.Until(ExpectedConditions.ElementExists(By.XPath(XPathStr)));
                wait.Until(ExpectedConditions.ElementExists(By.XPath(GuiTestParameter.ElementXpath)));

                //------------------------------------------------------//
                //-------------------- RADIO BUTTON --------------------//
                //------------------------------------------------------//
                if (GuiTestParameter.ElementType.CompareTo("RADIO_BUTTON") == 0 || GuiTestParameter.ElementType.CompareTo("BUTTON") == 0)
                {
                    #region RADIO BUTTON
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    //elementText = query.Text;
                    //MessageBox.Show("query.Text:" + query.Text);
                    query.Click();
                    #endregion
                }

                //------------------------------------------------------//
                //---------------------- CHECK BOX ---------------------//
                //------------------------------------------------------//
                else if (GuiTestParameter.ElementType.CompareTo("CHECK_BOX") == 0)
                {
                    #region CHECK BOX
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    if (GuiTestParameter.WriteValue.CompareTo("Y") == 0 && query.Selected == false)
                    {
                        query.Click();
                    }
                    else if (GuiTestParameter.WriteValue.CompareTo("N") == 0 && query.Selected == true)
                    {
                        query.Click();
                    }
                    #endregion
                }

                //------------------------------------------------------//
                //---------------------- TEXT BOX ----------------------//
                //------------------------------------------------------//
                else if (GuiTestParameter.ElementType.CompareTo("TEXT_BOX") == 0)
                {
                    #region TEXT BOX
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    query.Clear();
                    query.SendKeys(GuiTestParameter.WriteValue);
                    #endregion
                }

                //------------------------------------------------------//
                //------------------- DROP DOWN LIST -------------------//
                //------------------------------------------------------//
                else if (GuiTestParameter.ElementType.CompareTo("DROP_DOWN_LIST") == 0)
                {
                    #region DROP DOWN LIST
                    // 取得下拉選單元素
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    SelectElement select = new SelectElement(query);

                    // 選取下拉選單項目
                    int seIndex = Convert.ToInt32(GuiTestParameter.WriteValue);
                    select.SelectByIndex(seIndex);

                    // 取得被選取的字串
                    Thread.Sleep(2000);
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    SelectElement newSelect = new SelectElement(query); // 重新 attach element to browser
                    GuiTestParameter.WriteValue = newSelect.SelectedOption.Text;
                    #endregion
                }

            }
            catch (Exception ex)
            {
                GuiTestParameter.Note = "SetCGA2121ElementValue() Error: \n" + ex.ToString();
                return false;
            }
            Thread.Sleep(2000);
            return true;
        }

        public bool GetWebElementValue(ref GuiScriptParameter GuiTestParameter)
        {
            IWebElement query;
            bool Result = true;

            //Thread.Sleep(2000);
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                wait.Until(ExpectedConditions.ElementExists(By.XPath(GuiTestParameter.ElementXpath)));


                //------------------------------------------------------//
                //-------------------- RADIO BUTTON --------------------//
                //------------------------------------------------------//
                if (GuiTestParameter.ElementType.CompareTo("RADIO_BUTTON") == 0)
                {
                    #region RADIO BUTTON
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    //GetElementText = query.Text;
                    Result = query.Selected;
                    if (Result == true)
                        GuiTestParameter.GetValue = GuiTestParameter.ExpectedValue;
                    #endregion
                }

                //------------------------------------------------------//
                //---------------------- CHECK BOX ---------------------//
                //------------------------------------------------------//
                else if (GuiTestParameter.ElementType.CompareTo("CHECK_BOX") == 0)
                {
                    #region CHECK BOX
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    if (GuiTestParameter.ExpectedValue.CompareTo("Y") == 0)
                    {
                        if (query.Selected == true)
                        {
                            Result = true;
                            GuiTestParameter.GetValue = "Y";
                        }
                        else
                        {
                            Result = false;
                            GuiTestParameter.GetValue = "N";
                        }
                    }
                    else if (GuiTestParameter.ExpectedValue.CompareTo("N") == 0)
                    {
                        if (query.Selected == false)
                        {
                            Result = true;
                            GuiTestParameter.GetValue = "N";
                        }
                        else
                        {
                            Result = false;
                            GuiTestParameter.GetValue = "Y";
                        }
                    }
                    #endregion
                }

                //------------------------------------------------------//
                //---------------------- TEXT BOX ----------------------//
                //------------------------------------------------------//
                else if (GuiTestParameter.ElementType.CompareTo("TEXT_BOX") == 0)
                {
                    #region TEXT BOX
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    GuiTestParameter.GetValue = query.GetAttribute("value");

                    if (GuiTestParameter.GetValue.CompareTo(GuiTestParameter.ExpectedValue) == 0)
                        Result = true;
                    else
                        Result = false;
                    #endregion
                }

                //------------------------------------------------------//
                //------------------- DROP DOWN LIST -------------------//
                //------------------------------------------------------//
                else if (GuiTestParameter.ElementType.CompareTo("DROP_DOWN_LIST") == 0)
                {
                    #region DROP DOWN LIST
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    SelectElement select = new SelectElement(query);
                    GuiTestParameter.GetValue = select.SelectedOption.Text;

                    // 選取下拉選單項目
                    int seIndex = Convert.ToInt32(GuiTestParameter.ExpectedValue);
                    select.SelectByIndex(seIndex);

                    // 取得被選取的字串
                    Thread.Sleep(2000);
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    SelectElement newSelect = new SelectElement(query); // 重新 attach element to browser
                    GuiTestParameter.ExpectedValue = newSelect.SelectedOption.Text;
                    //MessageBox.Show("selectedText:" + elementText);

                    if (GuiTestParameter.GetValue.CompareTo(GuiTestParameter.ExpectedValue) == 0)
                        Result = true;
                    else
                        Result = false;
                    #endregion
                }

                //------------------------------------------------------//
                //------------------------ TABLE -----------------------//
                //------------------------------------------------------//
                else if (GuiTestParameter.ElementType.CompareTo("TABLE") == 0)
                {
                    #region TABLE
                    query = driver.FindElement(By.XPath(GuiTestParameter.ElementXpath));
                    GuiTestParameter.GetValue = query.Text;

                    if (GuiTestParameter.GetValue.CompareTo(GuiTestParameter.ExpectedValue) == 0)
                        Result = true;
                    else
                        Result = false;
                    #endregion
                }
            }
            catch (Exception ex)
            {
                GuiTestParameter.Note = "GetCGA2121ElementValue() Error: \n" + ex.ToString();
                return false;
            }
            return Result;
        }

        public bool GetBrowserVersion(ref string BrowserVer)
        {
            try
            {
                BrowserVer = BrowserCapabilities.Version;
            }
            catch
            {
                return false;
            }

            return true;
        }
        
        public bool GoToURL(string strURL)
        {
            try
            {
                driver.Navigate().GoToUrl(strURL);
                Thread.Sleep(2000);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool SettingTimeout(int timeLimit)  //** Setting Timeout (second)
        {
            try
            {
                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(timeLimit);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool GetCurrentURL(ref string currentURL)
        {
            try
            {
                currentURL = driver.Url;
                Thread.Sleep(2000);
            }
            catch
            {
                return false;
            }

            return true;
        }
        
        public bool RefreshWebPage()               //** Refresh Web page
        {
            try
            {
                driver.Navigate().Refresh();
                Thread.Sleep(2000);
            }
            catch
            {
                return false;
            }
            return true;
        }

        public bool WindowMaximize()
        {
            try
            {
                if (TestBrowser.CompareTo("firefox") != 0)
                    driver.Manage().Window.Maximize();
            }
            catch
            {
                return false;
            }

            return true;
        }

        public void CheckAlertMessage()
        {
            for (int i = 0; i < 2; i++)
            {
                try
                {
                    IAlert alert = driver.SwitchTo().Alert(); // Check the presence of alert
                    //MessageBox.Show("alert.Text: " + alert.Text);
                    alert.Accept();     // Accept consume the alert
                    Thread.Sleep(6000);
                }
                catch (NoAlertPresentException ex) // Alert not present
                {
                    Debug.WriteLine("Alert not present: " + ex.ToString());
                    break;
                }

                Thread.Sleep(6000);
            }
        }

        private bool isElementPresent()
        {
            try
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
                //driver.findElement(by);
                return true;

            }
            catch (NoSuchElementException e)
            {
                return false;
            }
        }

        public bool Close_WebDriver()         //** Quits this selenium driver, closing every associated window.
        {
            try
            {
                //driver.Quit();
                //Thread.Sleep(2000);

                //** Remove the Browser process
                System.Diagnostics.Process[] browserProc = System.Diagnostics.Process.GetProcessesByName(TestBrowser);
                if (browserProc.Length > 0)
                {
                    for (int i = 0; i < browserProc.Length; i++)
                        browserProc[i].Kill();
                }

                Thread.Sleep(1000);

                //** Remove the Browser Driver process
                System.Diagnostics.Process[] driverProc = System.Diagnostics.Process.GetProcessesByName(BrowserDriverProcessesName);
                if (driverProc.Length > 0)
                {
                    for (int i = 0; i < driverProc.Length; i++)
                        driverProc[i].Kill();
                }
            }
            catch
            {
                return false;
            }
            return true;
        }






#region For Common Case Script (Test Common Case GUI)
       
        public bool SetCGA2121ElementValue(string ElementType, string ElementXpath, ref string WriteValue, ref string ExceptionInfo)
        {
            IWebElement query;

            //Thread.Sleep(2000);
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30)); //.Until(ExpectedConditions.ElementExists(By.XPath(XPathStr)));
                wait.Until(ExpectedConditions.ElementExists(By.XPath(ElementXpath)));

                if (ElementType.CompareTo("RADIO_BUTTON") == 0 || ElementType.CompareTo("BUTTON") == 0)
                {
                    query = driver.FindElement(By.XPath(ElementXpath));
                    //elementText = query.Text;
                    //MessageBox.Show("query.Text:" + query.Text);
                    query.Click();
                }
                else if (ElementType.CompareTo("CHECK_BOX") == 0)
                {
                    query = driver.FindElement(By.XPath(ElementXpath));
                    if (WriteValue.CompareTo("Y") == 0 && query.Selected == false)
                    {
                        query.Click();
                    }
                    else if (WriteValue.CompareTo("N") == 0 && query.Selected == true)
                    {
                        query.Click();
                    }
                }
                else if (ElementType.CompareTo("TEXT_BOX") == 0)
                {
                    query = driver.FindElement(By.XPath(ElementXpath));
                    query.Clear();
                    query.SendKeys(WriteValue);
                }
                else if (ElementType.CompareTo("DROP_DOWN_LIST") == 0)
                {
                    // 取得下拉選單元素
                    query = driver.FindElement(By.XPath(ElementXpath));
                    SelectElement select = new SelectElement(query);

                    // 選取下拉選單項目
                    int seIndex = Convert.ToInt32(WriteValue);
                    select.SelectByIndex(seIndex);

                    // 取得被選取的字串
                    Thread.Sleep(2000);
                    query = driver.FindElement(By.XPath(ElementXpath));
                    SelectElement newSelect = new SelectElement(query); // 重新 attach element to browser
                    WriteValue = newSelect.SelectedOption.Text;
                }

            }
            catch (Exception ex)
            {
                ExceptionInfo = "SetCGA2121ElementValue() Error: \n" + ex.ToString();
                return false;
            }
            Thread.Sleep(2000);

            return true;
        }

        public bool GetCGA2121ElementValue(string ElementType, string ElementXpath, ref string ExpectedValue, ref string GetValue, ref string ExceptionInfo)
        {
            IWebElement query;
            bool Result = true;

            //Thread.Sleep(2000);
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                wait.Until(ExpectedConditions.ElementExists(By.XPath(ElementXpath)));

                if (ElementType.CompareTo("RADIO_BUTTON") == 0)
                {
                    query = driver.FindElement(By.XPath(ElementXpath));
                    //GetElementText = query.Text;
                    Result = query.Selected;
                    if (Result == true)
                        GetValue = ExpectedValue;
                }
                else if (ElementType.CompareTo("CHECK_BOX") == 0)
                {
                    query = driver.FindElement(By.XPath(ElementXpath));
                    if (ExpectedValue.CompareTo("Y") == 0)
                    {
                        if (query.Selected == true)
                        {
                            Result = true;
                            GetValue = "Y";
                        }
                        else
                        {
                            Result = false;
                            GetValue = "N";
                        }
                    }
                    else if (ExpectedValue.CompareTo("N") == 0)
                    {
                        if (query.Selected == false)
                        {
                            Result = true;
                            GetValue = "N";
                        }
                        else
                        {
                            Result = false;
                            GetValue = "Y";
                        }
                    }
                }
                else if (ElementType.CompareTo("TEXT_BOX") == 0)
                {
                    query = driver.FindElement(By.XPath(ElementXpath));
                    GetValue = query.GetAttribute("value");

                    if (GetValue.CompareTo(ExpectedValue) == 0)
                        Result = true;
                    else
                        Result = false;
                }
                else if (ElementType.CompareTo("DROP_DOWN_LIST") == 0)
                {
                    query = driver.FindElement(By.XPath(ElementXpath));
                    SelectElement select = new SelectElement(query);
                    GetValue = select.SelectedOption.Text;

                    // 選取下拉選單項目
                    int seIndex = Convert.ToInt32(ExpectedValue);
                    select.SelectByIndex(seIndex);

                    // 取得被選取的字串
                    Thread.Sleep(2000);
                    query = driver.FindElement(By.XPath(ElementXpath));
                    SelectElement newSelect = new SelectElement(query); // 重新 attach element to browser
                    ExpectedValue = newSelect.SelectedOption.Text;
                    //MessageBox.Show("selectedText:" + elementText);

                    if (GetValue.CompareTo(ExpectedValue) == 0)
                        Result = true;
                    else
                        Result = false;
                }
                else if (ElementType.CompareTo("TABLE") == 0)
                {
                    query = driver.FindElement(By.XPath(ElementXpath));
                    GetValue = query.Text;

                    if (GetValue.CompareTo(ExpectedValue) == 0)
                        Result = true;
                    else
                        Result = false;
                }
            }
            catch (Exception ex)
            {
                ExceptionInfo = "GetCGA2121ElementValue() Error: \n" + ex.ToString();
                return false;
            }

            return Result;
        } 

        public bool CheckPageException(ref string ExceptionInfo)
        {
            //IWebElement query;
            string ReloadXPath = "/html/body/h1/a";  // RELOAD
            bool bPageException = false;

            try
            {
                Thread.Sleep(1000);
                bPageException = driver.FindElement(By.XPath(ReloadXPath)).Displayed;
            }
            catch (Exception ex)
            {
            }

            return bPageException;
        }

        
#endregion



    } //--- End of class SeleniumAPI  
}

/* CBT_SeleniumApi Eample:
 *
 * using NS_CbtSnmpClass
 * 
 * CBT_SeleniumApi cs_BROWSER = new CBT_SeleniumApi();
 * // Initialize; TestBrowser keyword: "chrome" / "iexplore" / "firefox"
 * cs_BROWSER.init(TestBrowser);
 * // Go to URL
 * GoToCGA2121_URL(URLstr);
 * // Set time limit for web driver action
 * SettingTimeout(timeLimit);
 * // Set Web browser window to maximize
 * WindowMaximize();
 * // Set Value
 * SetCGA2121ElementValue(ElementType, ElementXpath, ref WriteValue, ref ExceptionInfo);
 * // Get Value
 * GetCGA2121ElementValue(ElementType, ElementXpath, ref ExpectedValue, ref GetValue, ref ExceptionInfo);
 * // Refresh Web page
 * RefreshCGA2121();
 * // Close WebDriver
 * Close_WebDriver();
 */