///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RvRFunctionTest.cs
///  Update         : 2015-05-11
///  Description    : Main function
///  Modified       : 2015-05-11 Initial version  
///                   
///---------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;
using AgilentInstruments ;
using Ivi.Visa.Interop;
//using NationalInstruments.VisaNS;
using System.Xml;
using System.Net.Sockets;
using System.Net;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {     
        bool bFunctionTestRunning = false;
        string m_RvRTestSubFolder;
       
        Thread threadRvRFunctionTest;
        
        
        /* Declare Test Condition related variable */
        string[,] testConfig; //test condition data
        bool bTest_2_4G = false;
        bool bTest_5G = false;
        
        

        bool bCheckPoint;
        string WiFiData;
        StringBuilder linkrate_csv;
        /* Declare delegate function */
        public delegate void showFunctionTestGUIDelegate();

        private void rvRTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sTestItem = "RvR Test";
            ToggleToolStripMenuItem(false);
            tabControl_RouterStartPage.Hide();

            rvRTestToolStripMenuItem.Checked = true;

            

            //foreach (TabPage page in RvRTest_tabControl.TabPages)
            //{
            //    if (page.Text.Equals("RvR Test Condition", StringComparison.OrdinalIgnoreCase) || page.Text.Equals("RvR Attenuator Setting", StringComparison.OrdinalIgnoreCase) || page.Text.Equals("RvR Function Test", StringComparison.OrdinalIgnoreCase))
            //        continue;

            //    RvRTest_tabControl.TabPages.Remove(page);
            //}

            RvRTabHideAllPages();

            tp_RvRTestCondition.Parent = this.RvRTest_tabControl; //Show tabPage
            tp_RvRAttenuatorSetting.Parent = this.RvRTest_tabControl;
            tp_RvRFunctionTest.Parent = this.RvRTest_tabControl;

            RvRTest_tabControl.Show();
            tsslMessage.Text = RvRTest_tabControl.TabPages[RvRTest_tabControl.SelectedIndex].Text + " Control Panel";

            string xmlFile = System.Windows.Forms.Application.StartupPath + "\\config\\" + Constants.XML_ROUTER_RvR;
            Debug.WriteLine(File.Exists(xmlFile) ? "File exist" : "File is not existing");

            if (!File.Exists(xmlFile))
            {
                WriteXmlDefaultRvRTest(xmlFile);
            }

            ReadXmlRvRTest(xmlFile);

            InitRvRTestCondition();
        }

        private void tabControl_RvRTest_Selected(object sender, TabControlEventArgs e)
        {
            if(RvRTest_tabControl.SelectedIndex >=0)
                tsslMessage.Text = RvRTest_tabControl.TabPages[RvRTest_tabControl.SelectedIndex].Text + " Control Panel";
        }

        private void txt_ChariotSetting_ChariotFolder_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.DefaultExt = "exe";
            openFileDialog1.Title = "Choose Chariot.exe file ";
            openFileDialog1.Filter = "exe files(*.exe)|*.exe|All files(*.*)|*.*";
            openFileDialog1.FileName = "IxChariot.exe";
            openFileDialog1.FilterIndex = 1;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtRvRTurnFunctionTestChariotFolder.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open Chariot.exe: " + ex.Message);
                }
            }
        }

        private void btnRvRFunctionTestRun_Click(object sender, EventArgs e)
        {       
            WiFiData = "";
            linkrate_csv = new StringBuilder();
            /* Prevent double-click from double firing the same action */
            btnRvRFunctionTestRun.Enabled = false;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            btnRvRFunctionTestRun.Text = "Stop" ;            

            if (bFunctionTestRunning == false)
            {
                bCheckPoint = false;
                if (chkbox_RouterSetting_ConfigRouterManually.Checked == true)
                    bConfigRouter = false;
                else bConfigRouter = true;


                if (threadRvRFunctionTest != null) 
                    threadRvRFunctionTest.Abort();  
              
                

                /* Check if runtst.exe and fmttst.exe exist */
                string chariotFile = txt_ChariotSetting_ChariotFolder.Text;
                string tstFile = chariotFile.Substring(0, chariotFile.Length - Path.GetFileName(chariotFile).Length);
                Path_runtst = tstFile + "runtst.exe";
                Path_fmttst = tstFile + "fmttst.exe";

                SetText("Checking runtst.exe and fmttst.ext whether exist or not.");
                if (!File.Exists(Path_runtst) || !File.Exists(Path_fmttst))
                {
                    MessageBox.Show("The runtst.exe or the fmttst.exe file doesn't exist, please specify a chariot.exe in which folder!!!","Warning");
                    btnRvRFunctionTestRun.Text = "Run" ;
                    btnRvRFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;                    
                    return;
                }
                
                /* Config and Check Attenuator Setting */
                //SetText("Checking the attenuator settings whether a correct value.");
                if (!RvRCheckCondition())
                {
                    MessageBox.Show("Condition check failed. Please check information for detailed!!!", "Warning");
                    btnRvRFunctionTestRun.Text = "Run" ;
                    btnRvRFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;                    
                    return;
                }                

                /* Check IP address is valid */
                SetText("Checking IP address whether in a valid.");
                if(!CheckIPValid(txt_RouterSetting_DUTIPAddress.Text) || 
                   !CheckIPValid(txt_ClientSetting_24gIPaddress.Text) || 
                   !CheckIPValid(txt_ClientSetting_5gIPaddress.Text)) 
                {
                    MessageBox.Show("IP address is not valid. Please check your network settings!!!", "Warning");
                    btnRvRFunctionTestRun.Text = "Run" ;
                    btnRvRFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;                     
                    return;
                }

                if (bConfigRouter)
                {
                    /* Ping DUT */
                    SetText("Checking DUT whether in active mode.");
                    if (!PingClient(txt_RouterSetting_DUTIPAddress.Text, PingTimeout))
                    {
                        MessageBox.Show("Ping DUT failed. Please check your network environment!!!", "Warning");
                        btnRvRFunctionTestRun.Text = "Run";
                        btnRvRFunctionTestRun.Enabled = true;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        return;
                    }

                    /* Send pre-cgi command CloseSession and mfgTest=1 to DUT */
                    if (cbox_Model_Name.SelectedItem.ToString() == "E8350")
                    {
                        SetText("Checking AP router whether in MFG mode.");
                        if (!RunPreCGI(cbox_Model_Name.SelectedItem.ToString(), txt_RouterSetting_DUTIPAddress.Text, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text))
                        {
                            btnRvRFunctionTestRun.Text = "Run";
                            btnRvRFunctionTestRun.Enabled = true;
                            System.Windows.Forms.Cursor.Current = Cursors.Default;
                            return;
                        }
                    }


                    if (model == "EA8500X8")
                    {
                        security = cbox_TestCondition_Security.SelectedItem.ToString();
                        passphrase = txt_TestCondition_Passphrase.Text;
                        if (security != "None" && passphrase == "")
                        {
                            MessageBox.Show("Setting Passphras for security!!!", "Warning");
                            btnRvRFunctionTestRun.Text = "Run";
                            btnRvRFunctionTestRun.Enabled = true;
                            System.Windows.Forms.Cursor.Current = Cursors.Default;
                            return;
                        }
                    }



                }
                else //bConfigRouter = false, Manually setting 
                {
                    model = txt_Temp_ModelName.Text;
                }
                /* Check if attenuator trun on  */
                SetText("Checking attenuator device whether in active mode.");
                if (!CheckAtteuatorOn())
                {
                    MessageBox.Show("Attenuator is not ready. Please turn on the test equipment!!!", "Warning");
                    btnRvRFunctionTestRun.Text = "Run";
                    btnRvRFunctionTestRun.Enabled = true;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    return;
                }

                /* Open needed GPIB resource */
                SetText("OpenGPIBPort.");
                OpenGPIBport();
                                
                /* Read attenuator value for 2.4G and 5G */
                SetText("Read Attenuator Value.");
                ReadAtteunatorValue();

                /* Create sub folder for saving the report data */
                SetText("Create folder.");
                if(chkbox_RouterSetting_ConfigRouterManually.Checked)
                    m_RvRTestSubFolder = createRvRTestSubFolder(txt_Temp_ModelName.Text);  
                else
                    m_RvRTestSubFolder = createRvRTestSubFolder(cbox_Model_Name.SelectedItem.ToString());                
                bFunctionTestRunning = true;
                /* Disable all controller */
                ToggleRvRFunctionTestControl(false);

                bCheckPoint = true;                
                
                SetText("============================================");
                SetText("Start RvR testing...please wait a moment for doing the device configuration.");
                                
                threadRvRFunctionTest = new Thread(new ThreadStart(DoRvRFunctionTest));
                threadRvRFunctionTest.Name = "RvRTest";
                threadRvRFunctionTest.Start();
            }
            else
            {
                if (MessageBox.Show("Do you want to stop the test?", "Stop", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    if (threadRvRFunctionTest != null) 
                        threadRvRFunctionTest.Abort();
                }
                else
                    return;

                bFunctionTestRunning = false;
                btnRvRFunctionTestRun.Text = "Run";
                ToggleRvRFunctionTestControl(true);
            }
            
            /* Button release */
            Thread.Sleep(3000);
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            btnRvRFunctionTestRun.Enabled = true;
        }

        private void DoRvRFunctionTest()
        {
            string band;
            string mode;
            int channel;
            string ssid;
            int start;
            int stop;
            int steps;
            string txtst;
            string rxtst;
            string bitst;
            string str = string.Empty;
            string currentTime;

            band = testConfig[0, 8];

            for (int i = 0; i < testConfig.GetLength(0); i++)
            {
                if (bFunctionTestRunning == false)
                {
                    break;
                }

                band = testConfig[i, 0];
                mode = testConfig[i, 1];
                channel = Int32.Parse(testConfig[i, 2]);
                ssid = testConfig[i, 3];
                start = Int32.Parse(testConfig[i,4]);
                stop = Int32.Parse(testConfig[i, 5]);
                steps = Int32.Parse(testConfig[i, 6]);
                txtst = testConfig[i, 7];
                rxtst = testConfig[i, 8];
                bitst = testConfig[i, 9];

                /* Config Router */
                str = String.Format("============================================");
                Invoke(new SetTextCallBack(SetText), new object[] { str });

                

                if (bConfigRouter)
                {
                    str = String.Format("Configure AP router as modelName:{0}, band:{1}, mode:{2}, channel:{3}, ssid:{4}", model, band, mode, channel, ssid);
                    Invoke(new SetTextCallBack(SetText), new object[] { str });

                    if (!ConfigureRouter(model, band, mode, channel, ssid, security, passphrase))
                        continue;
                    str = String.Format("Configure AP router Succeed.");
                    Invoke(new SetTextCallBack(SetText), new object[] { str });
                }
                

                if(band == "2.4G")
                {                   
                    /* Initial attenuator value that depends on "Start Value" */
                    ConfigureAtteuatorValue(Convert.ToInt32(nud_TestCondition_AtteuationMin.Value.ToString()), a11713A_2_4G, Attenuation_buttonValue_2_4G);

                    str = "Ping 2.4G client IP and checking endpoint whether be activated...";
                    Invoke(new SetTextCallBack(SetText), new object[] { str });

                    int nPing = 0;
                    while(true)
                    {
                        if (PingClient(this.txt_ClientSetting_24gIPaddress.Text, PingTimeout))
                        {
                            str = "Ping succeed!!!";
                            Invoke(new SetTextCallBack(SetText), new object[] { str });     
                            break;
                        }

                        if (nPing++ == 60)
                        {
                            str = "Ping failed!!!";
                            Invoke(new SetTextCallBack(SetText), new object[] { str });
                            bCheckPoint = false;
                            break;
                        }
                        Thread.Sleep(1000);
                    }

                    for (int j = start; j <= stop; j += steps)
                    {
                        if (bFunctionTestRunning == false)
                        {
                            break;
                        }

                        string input = testConfig[i, 5];
                        string outputPath = System.Windows.Forms.Application.StartupPath + @"\report\" + m_RvRTestSubFolder;
                        /* file name is E8350_2_4G_20M_channel_1_attenuatot_1*/
                        /* E8350_TX_2_4G_11n_20M_Channel_" + i.ToString() + "dB_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss") */
                        string outputHeader = @"\" + model + "_2_4G_11N_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_";
                        string output = string.Empty;

                        str = "Configure attenuator to [" + j.ToString() + "dB]";
                        Invoke(new SetTextCallBack(SetText), new object[] { str });     
                        ConfigureAtteuatorValue(j, a11713A_2_4G, Attenuation_buttonValue_2_4G);

                        currentTime = DateTime.Now.ToString("yyyy_MMdd_HHmmss");

                        str = "Start to run Chariot";
                        Invoke(new SetTextCallBack(SetText), new object[] { str });

                        string path = System.Windows.Forms.Application.StartupPath + @"\report\" + m_RvRTestSubFolder + @"\LinkRate";
                        string resultFile = path + "\\LinkResult.csv";

                        Thread.Sleep(10000);
                        /* Run Tx */
                        if (testConfig[i, 7] != "")
                        {                            
                            str = "TX script file: " + testConfig[i,7] ;
                            Invoke(new SetTextCallBack(SetText), new object[] { str });
                            input = testConfig[i, 7];
                            output = outputPath + outputHeader + "Tx_" + Path.GetFileName(input) + "_" + currentTime;
                            
                            RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                            // To record remote site's wifi data
                            //WiFiData = GetWiFiData();
                            //string linkRate = model + "_2_4G_11N_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Tx_" + Path.GetFileName(input) + "_" + currentTime;
                            //WriteResultCSV(linkRate + ".csv", resultFile);
                        }
                        //Thread.Sleep(15000);
                        Thread.Sleep(10000);
                        /* Run RX */
                        if (testConfig[i, 8] != "")
                        {
                            str = "RX script file: " + testConfig[i, 8];
                            Invoke(new SetTextCallBack(SetText), new object[] { str });
                            input = testConfig[i, 8];
                            output = outputPath + outputHeader + "Rx_" + Path.GetFileName(input) + "_" + currentTime;

                            RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                            // To record remote site's wifi data
                            //WiFiData = GetWiFiData();
                            //string linkRate = model + "_2_4G_11N_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Rx_" + Path.GetFileName(input) + "_" + currentTime;
                            //WriteResultCSV(linkRate + ".csv", resultFile);
                        }
                        //Thread.Sleep(15000);
                        Thread.Sleep(10000);
                        /* Run bi-Direction */
                        if (testConfig[i, 9] != "")
                        {
                            str = "Bi-Direction script file: " + testConfig[i, 9];
                            Invoke(new SetTextCallBack(SetText), new object[] { str });
                            input = testConfig[i, 9];
                            output = outputPath + outputHeader + "Bi_" + Path.GetFileName(input) + "_" + currentTime;

                            RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                            //// To record remote site's wifi data
                            //WiFiData = GetWiFiData();
                            //string linkRate = model + "_2_4G_11N_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Bi_" + Path.GetFileName(input) + "_" + currentTime;
                            //WriteResultCSV(linkRate + ".csv", resultFile);
                        }

                        str = "Chariot Finished.";
                        Invoke(new SetTextCallBack(SetText), new object[] { str });                        
                    }
                }

                if(band == "5G")
                {                    
                    /* Initial attenuator value that depends on "Start Value" */
                    ConfigureAtteuatorValue(Convert.ToInt32(nud_TestCondition_AtteuationMin.Value.ToString()), a11713A_5G, Attenuation_buttonValue_5G);

                    str = "Ping 5G client IP and checking endpoint whether be activated...";
                    Invoke(new SetTextCallBack(SetText), new object[] { str });

                    int nPing = 0;
                    while(true)
                    {
                        if (PingClient(this.txt_ClientSetting_5gIPaddress.Text, PingTimeout))
                        {
                            str = "Ping succeed!!!";
                            Invoke(new SetTextCallBack(SetText), new object[] { str });
                            break;
                        }

                        if (nPing++ == 60)
                        {
                            str = "Ping failed!!!";
                            Invoke(new SetTextCallBack(SetText), new object[] { str });
                            bCheckPoint = false;
                            break;
                        }
                        Thread.Sleep(1000);
                    }

                    for (int j = start; j <= stop; j += steps)
                    {
                        if (bFunctionTestRunning == false)
                        {
                            break;
                        }

                        string input = testConfig[i, 5];
                        string outputPath = System.Windows.Forms.Application.StartupPath + @"\report\" + m_RvRTestSubFolder;
                        /* file name is E8350_2_4G_20M_channel_1_attenuatot_1*/
                        /* E8350_TX_2_4G_11n_20M_Channel_" + i.ToString() + "_" + DateTime.Now.ToString("yyyy_MMdd_HHmmss") */
                        string outputHeader = @"\" + model + "_5G_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_";
                        string output = string.Empty;

                        str = "Configure attenuator to [" + j.ToString() + "dB]";
                        Invoke(new SetTextCallBack(SetText), new object[] { str });
                        ConfigureAtteuatorValue(j, a11713A_5G, Attenuation_buttonValue_5G);

                        currentTime = DateTime.Now.ToString("yyyy_MMdd_HHmmss");

                        str = "Start to run Chariot";
                        Invoke(new SetTextCallBack(SetText), new object[] { str });
                        //RunChariotConsole(Path_runtst, Path_fmttst, input, output);                        

                        string path = System.Windows.Forms.Application.StartupPath + @"\report\" + m_RvRTestSubFolder + @"\LinkRate";
                        string resultFile = path + "\\LinkResult.csv";

                        Thread.Sleep(10000);
                        /* Run Tx */
                        if (testConfig[i, 7] != "")
                        {
                            str = "TX script file: " + testConfig[i, 7];
                            Invoke(new SetTextCallBack(SetText), new object[] { str });
                            input = testConfig[i, 7];
                            output = outputPath + outputHeader + "Tx_" + Path.GetFileName(input) + "_" + currentTime;

                            RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                            // To record remote site's wifi data
                            //WiFiData = GetWiFiData();
                            //string linkRate = model + "_5G_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Tx_" + Path.GetFileName(input) + "_" + currentTime;
                            //WriteResultCSV(linkRate + ".csv", resultFile);
                        }
                        //Thread.Sleep(15000);
                        Thread.Sleep(10000);
                        /* Run RX */
                        if (testConfig[i, 8] != "")
                        {
                            str = "RX script file: " + testConfig[i, 8];
                            Invoke(new SetTextCallBack(SetText), new object[] { str });
                            input = testConfig[i, 8];
                            output = outputPath + outputHeader + "Rx_" + Path.GetFileName(input) + "_" + currentTime;

                            RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                            // To record remote site's wifi data
                            //WiFiData = GetWiFiData();
                            //string linkRate = model + "_5G_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Rx_" + Path.GetFileName(input) + "_" + currentTime;
                            //WriteResultCSV(linkRate + ".csv", resultFile);
                        }
                        //Thread.Sleep(15000);
                        Thread.Sleep(10000);
                        /* Run Bi-Direction */
                        if (testConfig[i, 9] != "")
                        {
                            str = "Bi-Direction script file: " + testConfig[i, 9];
                            Invoke(new SetTextCallBack(SetText), new object[] { str });
                            input = testConfig[i, 9];
                            output = outputPath + outputHeader + "Bi_" + Path.GetFileName(input) + "_" + currentTime;

                            RunChariotConsole(Path_runtst, Path_fmttst, input, output);

                            // To record remote site's wifi data
                            //WiFiData = GetWiFiData();
                            //string linkRate = model + "_5G_" + mode + "_Channel_" + channel.ToString() + "_Attenuator_" + j.ToString() + "dB_" + "Bi_" + Path.GetFileName(input) + "_" + currentTime;
                            //WriteResultCSV(linkRate + ".csv", resultFile);
                        }

                        str = "Chariot Finished.";
                        Invoke(new SetTextCallBack(SetText), new object[] { str });                        
                    }                    
                }
            }

            if (bCheckPoint)
            {
                /* Test finish , write test result to File */
                str = "Test completed!!";
                Invoke(new SetTextCallBack(SetText), new object[] { str });
                
            }
            else
            {
                str = "Some errors was found during testing!!!";
                Invoke(new SetTextCallBack(SetText), new object[] { str });
            }

            /* Move pdf file to PDF folder*/

            //string Pdfpath = System.Windows.Forms.Application.StartupPath + @"\report\" + m_RvRTestSubFolder + @"\PDF";
            //string srcPath = System.Windows.Forms.Application.StartupPath + @"\report\";
            
            //foreach (string filename in Directory.GetFiles(srcPath))
            //{
            //    if (filename.Substring(filename.Length - 3, 3).ToLower() == "pdf")
            //    {
                    
            //    }
                    
            //}

            /* Save data to result.csv */
            str = "Save data to result.csv file.";
            Invoke(new SetTextCallBack(SetText), new object[] { str });
            string pathcsv = System.Windows.Forms.Application.StartupPath + @"\report\" + m_RvRTestSubFolder;
            string resultCsv = pathcsv + "\\Result.csv";

            if (File.Exists(resultCsv))
                File.Delete(resultCsv);

            string[] files = Directory.GetFiles(pathcsv);
            WriteCSV(files, resultCsv);



            bFunctionTestRunning = false;
            this.Invoke(new showFunctionTestGUIDelegate(ToggleFunctionTestRvR));
        }

        /* Check all needed Condition and Attenuator is ready*/
        private bool RvRCheckCondition()
        {
            /* Check if Test Condition Data is Empty */
            if (dgvRvRTestConditionData.RowCount <= 1)
            {
                MessageBox.Show("Config Test Condition Data Empty, Config First!!!");                
                return false;
            }

            /* Read Test Config and check if TX, RX, Bi tst file exists */
            if (!ReadTestConfig_RvR())
            {
                SetText("Read Test Config Failed.");
                return false ;
            }

            /* Check at least one GPIB number has set */
            if (nudRvRAtteuatorNumber24G.Value == 0 && nudRvRAtteuatorNumber5G.Value == 0)
            {
                MessageBox.Show("Both attenuator number is Zero. Please Check!!!", "Warning");
                return false;
            }

            /* Check if test band and test attenuator is match */
            if (bTest_2_4G)
            {
                if (nudRvRAtteuatorNumber24G.Value == 0)
                {
                    MessageBox.Show("Test Condition 2.4G needed!!!Set 2.4G Attenuator First!!!", "Warning");
                    return false;
                }
            }
            if (bTest_5G)
            {
                if (nudRvRAtteuatorNumber5G.Value == 0)
                {
                    MessageBox.Show("Test Condition 5G needed!!!Set 5G Attenuator First!!!", "Warning");
                    return false;
                }
            }

            /* Check if attenuator number and ip address number is the same */
            if(lboxRvRAtteuationSettingGPIBIP24G.Items.Count != nudRvRAtteuatorNumber24G.Value)
            {
                MessageBox.Show("2.4G: Attenuator number is not equal to the ip number!!!", "Warning");
                return false;
            }

            if(lboxRvRAtteuationSettingGPIBIP5G.Items.Count != nudRvRAtteuatorNumber5G.Value)
            {
                MessageBox.Show("5G: Attenuator number is not equal to the ip number!!!", "Warning");
                return false;
            }            
            return true;
        }

        private bool ReadTestConfig_RvR()
        {
            string[,] rowdata = new string[dgvRvRTestConditionData.RowCount - 1, 10] ;

            bTest_2_4G = false;
            bTest_5G = false;

            for (int i = 0; i < dgvRvRTestConditionData.RowCount-1; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    rowdata[i, j] = dgvRvRTestConditionData.Rows[i].Cells[j].Value.ToString();
                    /* for j=0 : Set bTest_2_4G=true  if need to test 2.4G , 
                     * bTest_5G = true  if need to test 5G */
                    if (j == 0)
                    {
                        if (rowdata[i, j] == "2.4G" && !bTest_2_4G) bTest_2_4G = true;
                        if (rowdata[i, j] == "5G" && !bTest_5G) bTest_5G = true;
                    }


                    /* For j = 7~9 : Check if tst File exist */
                    if (j >= 7 && j <= 9)
                    {
                        if ( rowdata[i,j]!="" && !File.Exists(rowdata[i, j]))
                        {
                            MessageBox.Show(rowdata[i,j] + " File doesn't exist. Please Check!!", "Warning" );
                            return false;
                        }
                    }
                }                
            }
            testConfig = rowdata;
            return true;
        }        

        private void ToggleRvRFunctionTestControl(bool Toggle)
        {
            /* Test Condition */
            rbtnRvRTestCondition24G.Enabled = Toggle;
            rbtnRvRTestCondition5G.Enabled = Toggle;
            cboxRvRTestConditionWirelessMode.Enabled = Toggle;
            cbox_Model_Name.Enabled = Toggle;
            cboxRvRTestConditionWirelessChannel.Enabled = Toggle;
            txt_TestCondition_SSID.Enabled = Toggle;
            nud_TestCondition_AtteuationMin.Enabled = Toggle;
            nud_TestCondition_AtteuationMax.Enabled = Toggle;
            nud_TestCondition_AtteuationStep.Enabled = Toggle;
            txtRvRTestConditionTxTst.Enabled = Toggle;
            txtRvRTestConditionRxTst.Enabled = Toggle;
            txtRvRTestConditionBiTst.Enabled = Toggle;
            btnRvRTestConditionAddSetting.Enabled = Toggle;
            btnRvRTestConditionEditSetting.Enabled = Toggle;
           
            /* Attenuator Setting */
            nudRvRAtteuation24GX1.Enabled = Toggle;
            nudRvRAtteuation24GX2.Enabled = Toggle;
            nudRvRAtteuation24GX3.Enabled = Toggle;
            nudRvRAtteuation24GX4.Enabled = Toggle;
            nudRvRAtteuation24GX5.Enabled = Toggle;
            nudRvRAtteuation24GX6.Enabled = Toggle;
            nudRvRAtteuation24GX7.Enabled = Toggle;
            nudRvRAtteuation24GX8.Enabled = Toggle;            
            nudRvRGPIBInterface24G.Enabled = Toggle;
            nudRvRAtteuatorNumber24G.Enabled = Toggle;
            nudRvRAtteuationGPIBIPAdderss24G.Enabled = Toggle;
            labRvRAttenuationAdd24G.Enabled = Toggle;
            labRvRAttenuationDel24G.Enabled = Toggle;
            lboxRvRAtteuationSettingGPIBIP24G.Enabled = Toggle;           

            nudRvRAtteuation5GX1.Enabled = Toggle;
            nudRvRAtteuation5GX2.Enabled = Toggle;
            nudRvRAtteuation5GX3.Enabled = Toggle;
            nudRvRAtteuation5GX4.Enabled = Toggle;
            nudRvRAtteuation5GX5.Enabled = Toggle;
            nudRvRAtteuation5GX6.Enabled = Toggle;
            nudRvRAtteuation5GX7.Enabled = Toggle;
            nudRvRAtteuation5GX8.Enabled = Toggle;            
            nudRvRGPIBInterface5G.Enabled = Toggle;
            nudRvRAtteuatorNumber5G.Enabled = Toggle;
            nudRvRAttenuationGPIBIPAddress5G.Enabled = Toggle;
            labRvRAttenuationAdd5G.Enabled = Toggle;
            labRvRAttenuationDel5G.Enabled = Toggle;
            lboxRvRAtteuationSettingGPIBIP5G.Enabled = Toggle;
            
            /* Function Test */
            cbox_Model_Name.Enabled = Toggle;
            txt_Model_SerialNumber.Enabled = Toggle;
            txt_Model_HWVersion.Enabled = Toggle;
            txt_Model_SWVersion.Enabled = Toggle;
            
            /* Commented by James - for EA8500 */
            //txt_RouterSetting_UserName.Enabled = Toggle;
            //txt_RouterSetting_Password.Enabled = Toggle;
            
            
            txt_RouterSetting_DUTIPAddress.Enabled = Toggle;
            txt_ClientSetting_24gIPaddress.Enabled = Toggle;
            txt_ClientSetting_5gIPaddress.Enabled = Toggle;
            txt_ChariotSetting_ChariotFolder.Enabled = Toggle;

            btnRvRFunctionTestRun.Text = Toggle ? "Run" : "Stop";
            Debug.WriteLine(btnRvRFunctionTestRun.Text);

            /* Disable/Enable "Items" and "Setup" menu item */
            itemToolStripMenuItem.Enabled = Toggle;
            setupToolStripMenuItem.Enabled = Toggle;

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

            if (bFunctionTestRunning == false)
            {
                MessageBox.Show(this, "Test complete!!!", "Information", MessageBoxButtons.OK);
            }
        }

        private string createRvRTestSubFolder(string ModelName)
        {
            string subFolder = ((ModelName == "") ? "E8350_" : ModelName + "_") +
                    DateTime.Now.ToString("yyyy_MMdd_HHmmss");


            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder);
            
            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\LinkRate"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\LinkRate");

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\PDF"))
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\report\" + subFolder + @"\PDF");
            
            return subFolder;
        }

        private string GetAtteuatorValue(int value, decimal[] attenuator_value)
        {
            decimal MaxValue = 0;
            decimal MinValue = attenuator_value[0];            
            decimal dValue = value;
                
             /* Check if value exceed attenuatior's limit */
            foreach (decimal d in attenuator_value)
            {
                MaxValue += d;
            }

            if (value > MaxValue || value < MinValue)
            {
                Debug.WriteLine("Attenuator out of range");
                Debug.WriteLine("Instrument status error: G01");
                return null;
            }

            string temp="";
            /* Measure the value rage */
            for (int i = 0; i < 8; i++)
            {
                if (dValue - attenuator_value[7-i] >= 0)
                {                   
                    temp = temp + (8 - i).ToString();
                    dValue -= attenuator_value[7 - i];
                }
            }
            return temp;            
        }

        private bool ConfigureAtteuatorValue(int value, Agilent11713A[] resource, decimal[] attenuator_value)
        {
            decimal MaxValue = 0;
            //decimal MinValue = attenuator_value[0];
            decimal MinValue = 0;
            decimal dValue = value;

            /* Check if value exceed attenuatior's limit */           
            foreach (decimal d in attenuator_value)
            {
                MaxValue += d;
            }

            if (value > MaxValue || value < MinValue)
            {
                Debug.WriteLine("Attenuator out of range");
                Debug.WriteLine("Instrument status error: G01");
                return false;
            }

            string temp = "";
            /* Measure the value rage */
            for (int i = 0; i < 8; i++)
            {
                if (dValue - attenuator_value[7 - i] >= 0)
                {
                    temp = temp + (8 - i).ToString();
                    dValue -= attenuator_value[7 - i];
                }
            }
            
            /* Preset Attenuator and Write value */
            for (int i = 0; i < resource.Length; i++)
            {
                if (resource[i] != null)
                {
                    resource[i].preset();
                    resource[i].configureAttenuator(temp);
                }
                else
                {
                    Debug.WriteLine("Attenuator resource could not be arranged.");
                }
            }
            return true;
        }

        private int[] ArrayStringToInt(string[] str)
        {
            int[] temp = new int[str.Length];
            int i =0 ;
            foreach (string s in str)
            {
                temp[i++] = Int32.Parse(s);
            }
            return temp;
        }

        private void WriteXmlDefaultRvRTest(string filename)
        {
            XmlWriterSettings setting = new XmlWriterSettings();
            setting.Indent = true; //指定縮排

            XmlWriter writer = XmlWriter.Create(filename, setting);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by RvR Test program.");
            writer.WriteStartElement("CybertanATE");
            writer.WriteAttributeString("Item", "RvR Test");

            /* Write Test Condition setting */

            /* Write Attenuator Setting */
            writer.WriteStartElement("Attenuator");

            /* 2.4 G setting */
            writer.WriteStartElement("Band_2_4G");

            /* Write Attenuator Value */
            writer.WriteStartElement("AttenuatorValue_2_4G");
            writer.WriteElementString("X1", "1");
            writer.WriteElementString("X2", "2");
            writer.WriteElementString("X3", "4");
            writer.WriteElementString("X4", "4");
            writer.WriteElementString("X5", "10");
            writer.WriteElementString("X6", "20");
            writer.WriteElementString("X7", "30");
            writer.WriteElementString("X8", "30");
            writer.WriteEndElement();

            /* Write GPIB Interface*/
            writer.WriteStartElement("GPIB_Interface_2_4G");
            writer.WriteElementString("GPIB_Interface", "0");
            writer.WriteEndElement();

            /* Write GPIB Number */
            writer.WriteStartElement("GPIB_Number_2_4G");
            writer.WriteElementString("GPIB_Number", "0");
            writer.WriteEndElement();

            /*Write GPIB IP address*/
            writer.WriteStartElement("GPIB_IP_2_4G");
            
            writer.WriteEndElement();
            /* Writer */


            writer.WriteEndElement();
            //End of write 2.4G

            /* 5G Setting*/
            writer.WriteStartElement("Band_5G");

            /* Write Attenuator Value */
            writer.WriteStartElement("AttenuatorValue_5G");
            writer.WriteElementString("X1", "1");
            writer.WriteElementString("X2", "2");
            writer.WriteElementString("X3", "4");
            writer.WriteElementString("X4", "4");
            writer.WriteElementString("X5", "10");
            writer.WriteElementString("X6", "20");
            writer.WriteElementString("X7", "40");
            writer.WriteElementString("X8", "40");
            writer.WriteEndElement();

            /* Write GPIB Interface*/
            writer.WriteStartElement("GPIB_Interface_5G");
            writer.WriteElementString("GPIB_Interface", "0");
            writer.WriteEndElement();

            /* Write GPIB Number */
            writer.WriteStartElement("GPIB_Number_5G");
            writer.WriteElementString("GPIB_Number", "0");
            writer.WriteEndElement();

            /*Write GPIB IP address*/
            writer.WriteStartElement("GPIB_IP_5G");
            
            writer.WriteEndElement();

            writer.WriteEndElement();
            //End of write 5G

            writer.WriteEndElement();
            //End of write Attenuator


            /* Write Function Test Setting */
            writer.WriteStartElement("FunctionTest");
            // Model section
            writer.WriteStartElement("Model");
            writer.WriteElementString("Name", "EA8500");
            writer.WriteElementString("SN", "123456789");
            writer.WriteElementString("SWVer", "1.0");
            writer.WriteElementString("HWVer", "1.0");
            writer.WriteEndElement();

            // Router Setting
            writer.WriteStartElement("Router");
            writer.WriteElementString("Username", "admin");
            writer.WriteElementString("Password", "admin");
            writer.WriteElementString("IP", "192.168.1.1");
            writer.WriteEndElement();

            // Client Setting
            writer.WriteStartElement("Client");
            writer.WriteElementString("IPaddr_2_4G", "192.168.1.100");
            writer.WriteElementString("IPaddr_5G", "192.168.1.101");
            writer.WriteEndElement();

            //ChariotSetting
            writer.WriteStartElement("Chariot");
            writer.WriteElementString("File", @"C:\Program Files (x86)\Ixia\IxChariot\IxChariot.exe");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }

        private void WriteXmlRvRTest(string filename)
        {
            XmlWriterSettings setting = new XmlWriterSettings();
            setting.Indent = true; //指定縮排

            XmlWriter writer = XmlWriter.Create(filename, setting);
            writer.WriteStartDocument();
            writer.WriteComment("DO NOT MODIFY THIS FILE. This file was generated by RvR Test program.");
            writer.WriteStartElement("CybertanATE") ;
            writer.WriteAttributeString("Item","RvR Test") ;

            /* Write Test Condition setting */

            /* Write Attenuator Setting */
            writer.WriteStartElement("Attenuator") ;

            /* 2.4 G setting */
            writer.WriteStartElement("Band_2_4G") ;

            /* Write Attenuator Value */
            writer.WriteStartElement("AttenuatorValue_2_4G") ;
            writer.WriteElementString("X1", nudRvRAtteuation24GX1.Value.ToString()) ;
            writer.WriteElementString("X2", nudRvRAtteuation24GX2.Value.ToString()) ;
            writer.WriteElementString("X3", nudRvRAtteuation24GX3.Value.ToString()) ;
            writer.WriteElementString("X4", nudRvRAtteuation24GX4.Value.ToString()) ;
            writer.WriteElementString("X5", nudRvRAtteuation24GX5.Value.ToString()) ;
            writer.WriteElementString("X6", nudRvRAtteuation24GX6.Value.ToString()) ;
            writer.WriteElementString("X7", nudRvRAtteuation24GX7.Value.ToString()) ;
            writer.WriteElementString("X8", nudRvRAtteuation24GX8.Value.ToString()) ;
            writer.WriteEndElement();
            
            /* Write GPIB Interface*/
            writer.WriteStartElement("GPIB_Interface_2_4G");
            writer.WriteElementString("GPIB_Interface", nudRvRGPIBInterface24G.Value.ToString());
            writer.WriteEndElement() ;

            /* Write GPIB Number */
            writer.WriteStartElement("GPIB_Number_2_4G") ;
            writer.WriteElementString("GPIB_Number", nudRvRAtteuatorNumber24G.Value.ToString()) ;
            writer.WriteEndElement();

            /*Write GPIB IP address*/
            writer.WriteStartElement("GPIB_IP_2_4G") ;
            int i=1 ; 
            foreach(string s in lboxRvRAtteuationSettingGPIBIP24G.Items)
            {
                writer.WriteElementString("IP_" + i.ToString(), s) ;
                i++;
            }                
            writer.WriteEndElement();
            /* Writer */


            writer.WriteEndElement();
            //End of write 2.4G

            /* 5G Setting*/
            writer.WriteStartElement("Band_5G") ;

            /* Write Attenuator Value */
            writer.WriteStartElement("AttenuatorValue_5G") ;
            writer.WriteElementString("X1", nudRvRAtteuation5GX1.Value.ToString()) ;
            writer.WriteElementString("X2", nudRvRAtteuation5GX2.Value.ToString()) ;
            writer.WriteElementString("X3", nudRvRAtteuation5GX3.Value.ToString()) ;
            writer.WriteElementString("X4", nudRvRAtteuation5GX4.Value.ToString()) ;
            writer.WriteElementString("X5", nudRvRAtteuation5GX5.Value.ToString()) ;
            writer.WriteElementString("X6", nudRvRAtteuation5GX6.Value.ToString()) ;
            writer.WriteElementString("X7", nudRvRAtteuation5GX7.Value.ToString()) ;
            writer.WriteElementString("X8", nudRvRAtteuation5GX8.Value.ToString()) ;
            writer.WriteEndElement();

            /* Write GPIB Interface*/
            writer.WriteStartElement("GPIB_Interface_5G");
            writer.WriteElementString("GPIB_Interface", nudRvRGPIBInterface5G.Value.ToString());
            writer.WriteEndElement() ;

            /* Write GPIB Number */
            writer.WriteStartElement("GPIB_Number_5G") ;
            writer.WriteElementString("GPIB_Number", nudRvRAtteuatorNumber5G.Value.ToString()) ;
            writer.WriteEndElement();

            /*Write GPIB IP address*/
            writer.WriteStartElement("GPIB_IP_5G") ;
            i = 1;  
            foreach(string s in lboxRvRAtteuationSettingGPIBIP5G.Items)
            {
                writer.WriteElementString("IP_" + i.ToString(), s) ;
                i++;
            }                
            writer.WriteEndElement();

            writer.WriteEndElement();
            //End of write 5G

            writer.WriteEndElement() ;
            //End of write Attenuator


            /* Write Function Test Setting */
            writer.WriteStartElement("FunctionTest");
            // Model section
            writer.WriteStartElement("Model");
            writer.WriteElementString("Name", cbox_Model_Name.SelectedIndex.ToString());
            writer.WriteElementString("SN", txt_Model_SerialNumber.Text);
            writer.WriteElementString("SWVer", txt_Model_SWVersion.Text);
            writer.WriteElementString("HWVer", txt_Model_HWVersion.Text);
            writer.WriteEndElement();

            // Router Setting
            writer.WriteStartElement("Router");
            writer.WriteElementString("Username", txt_RouterSetting_UserName.Text);
            writer.WriteElementString("Password", txt_RouterSetting_Password.Text);
            writer.WriteElementString("IP", txt_RouterSetting_DUTIPAddress.Text);                
            writer.WriteEndElement();

            // Client Setting
            writer.WriteStartElement("Client");
            writer.WriteElementString("IPaddr_2_4G", txt_ClientSetting_24gIPaddress.Text);
            writer.WriteElementString("IPaddr_5G", txt_ClientSetting_5gIPaddress.Text);                
            writer.WriteEndElement();

            //ChariotSetting
            writer.WriteStartElement("Chariot");
            writer.WriteElementString("File", txt_ChariotSetting_ChariotFolder.Text);                
            writer.WriteEndElement();            
            writer.WriteEndElement();

            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }

        private bool ReadXmlRvRTest(string filename)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(filename);

            XmlNode node = doc.SelectSingleNode("CybertanATE");
            if (node == null)
            {
                return false;
            }

            XmlElement element = (XmlElement)node;
            string strID = element.GetAttribute("Item");
            Debug.Write(strID);
            if (strID.CompareTo("RvR Test") != 0)
            {
                MessageBox.Show("This XML file is incorrect", "Error");
                return false;
            }

            /* Read Attenuator Setting */
            /* Read  2.4G attenuator value x1-x8 */
            XmlNode nodeAttenuator = doc.SelectSingleNode("/CybertanATE/Attenuator/Band_2_4G/AttenuatorValue_2_4G");

            try
            {
                string x1 = nodeAttenuator.SelectSingleNode("X1").InnerText;
                string x2 = nodeAttenuator.SelectSingleNode("X2").InnerText;
                string x3 = nodeAttenuator.SelectSingleNode("X3").InnerText;
                string x4 = nodeAttenuator.SelectSingleNode("X4").InnerText;
                string x5 = nodeAttenuator.SelectSingleNode("X5").InnerText;
                string x6 = nodeAttenuator.SelectSingleNode("X6").InnerText;
                string x7 = nodeAttenuator.SelectSingleNode("X7").InnerText;
                string x8 = nodeAttenuator.SelectSingleNode("X8").InnerText;
                Debug.WriteLine("X1: " + x1);
                Debug.WriteLine("X2: " + x2);
                Debug.WriteLine("X3: " + x1);
                Debug.WriteLine("X4: " + x2);
                Debug.WriteLine("X5: " + x1);
                Debug.WriteLine("X6: " + x2);
                Debug.WriteLine("X7: " + x1);
                Debug.WriteLine("X8: " + x2);

                nudRvRAtteuation24GX1.Value = Decimal.Parse(x1);
                nudRvRAtteuation24GX2.Value = Decimal.Parse(x2);
                nudRvRAtteuation24GX3.Value = Decimal.Parse(x3);
                nudRvRAtteuation24GX4.Value = Decimal.Parse(x4);
                nudRvRAtteuation24GX5.Value = Decimal.Parse(x5);
                nudRvRAtteuation24GX6.Value = Decimal.Parse(x6);
                nudRvRAtteuation24GX7.Value = Decimal.Parse(x7);
                nudRvRAtteuation24GX8.Value = Decimal.Parse(x8);
            }
            catch(Exception ex)
            {
                Debug.WriteLine("/CybertanATE/Attenuator/Bnad_2_4G/Attenuator/" + ex);
            }
            
            /* Read 2.4G GPIB interface value */
            XmlNode nodeInterface_2_4G = doc.SelectSingleNode("/CybertanATE/Attenuator/Band_2_4G/GPIB_Interface_2_4G");
            try
            {
                string interface24 = nodeInterface_2_4G.SelectSingleNode("GPIB_Interface").InnerText;
                Debug.WriteLine("GPIB_Interface: " + interface24);
                nudRvRGPIBInterface24G.Value = Decimal.Parse(interface24);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/Attenuator/Band_2_4G/GPIB_Interface_2_4G "+ ex);
            }
            
            /* Read 2.4G GPIB number */
            XmlNode nodeNo_2_4G = doc.SelectSingleNode("/CybertanATE/Attenuator/Band_2_4G/GPIB_Number_2_4G");
            try
            {
                string no24 = nodeNo_2_4G.SelectSingleNode("GPIB_Number").InnerText;
                Debug.WriteLine("GPIB_Number: " + no24);

                nudRvRAtteuatorNumber24G.Value = Decimal.Parse(no24);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/Attenuator/Band_2_4G/GPIB_Number_2_4G " + ex);
            }

            /* Read 2.4G GPIB ip address */
            XmlNode nodeGpibBIP_2_4G = doc.SelectSingleNode("/CybertanATE/Attenuator/Band_2_4G/GPIB_IP_2_4G");
            /* Clear all items. Added by James */
            lboxRvRAtteuationSettingGPIBIP24G.Items.Clear();
            int i=1 ;
            for (i = 1; i < nudRvRAtteuatorNumber24G.Value+1; i++)
            {
                string temp = "IP_" + i.ToString();
                string ip = nodeGpibBIP_2_4G.SelectSingleNode("IP_"+i.ToString()).InnerText;
                Debug.WriteLine("IP_"+i.ToString()+" :" +ip);
                lboxRvRAtteuationSettingGPIBIP24G.Items.Add(ip);
            }

            /* Read 5G */
            /* Read 5G attenuator value x1-x8 */
            XmlNode nodeAttenuatorY = doc.SelectSingleNode("/CybertanATE/Attenuator/Band_5G/AttenuatorValue_5G");

            try
            {
                string x1 = nodeAttenuatorY.SelectSingleNode("X1").InnerText;
                string x2 = nodeAttenuatorY.SelectSingleNode("X2").InnerText;
                string x3 = nodeAttenuatorY.SelectSingleNode("X3").InnerText;
                string x4 = nodeAttenuatorY.SelectSingleNode("X4").InnerText;
                string x5 = nodeAttenuatorY.SelectSingleNode("X5").InnerText;
                string x6 = nodeAttenuatorY.SelectSingleNode("X6").InnerText;
                string x7 = nodeAttenuatorY.SelectSingleNode("X7").InnerText;
                string x8 = nodeAttenuatorY.SelectSingleNode("X8").InnerText;
                Debug.WriteLine("X1: " + x1);
                Debug.WriteLine("X2: " + x2);
                Debug.WriteLine("X3: " + x1);
                Debug.WriteLine("X4: " + x2);
                Debug.WriteLine("X5: " + x1);
                Debug.WriteLine("X6: " + x2);
                Debug.WriteLine("X7: " + x1);
                Debug.WriteLine("X8: " + x2);

                nudRvRAtteuation5GX1.Value = Decimal.Parse(x1);
                nudRvRAtteuation5GX2.Value = Decimal.Parse(x2);
                nudRvRAtteuation5GX3.Value = Decimal.Parse(x3);
                nudRvRAtteuation5GX4.Value = Decimal.Parse(x4);
                nudRvRAtteuation5GX5.Value = Decimal.Parse(x5);
                nudRvRAtteuation5GX6.Value = Decimal.Parse(x6);
                nudRvRAtteuation5GX7.Value = Decimal.Parse(x7);
                nudRvRAtteuation5GX8.Value = Decimal.Parse(x8);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/Attenuator/Bnad_5G/Attenuator/" + ex);
            }

            /* Read 5G GPIB interface value */
            XmlNode nodeInterface_5G = doc.SelectSingleNode("/CybertanATE/Attenuator/Band_5G/GPIB_Interface_5G");
            try
            {
                string interface5G = nodeInterface_5G.SelectSingleNode("GPIB_Interface").InnerText;
                Debug.WriteLine("GPIB_Interface: " + interface5G);
                nudRvRGPIBInterface5G.Value = Decimal.Parse(interface5G);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/Attenuator/Band_5G/GPIB_Interface_5G " + ex);
            }

            /* Read 5G GPIB number */
            XmlNode nodeNo_5G = doc.SelectSingleNode("/CybertanATE/Attenuator/Band_5G/GPIB_Number_5G");
            try
            {
                string no5G = nodeNo_5G.SelectSingleNode("GPIB_Number").InnerText;
                Debug.WriteLine("GPIB_Number: " + no5G);

                nudRvRAtteuatorNumber5G.Value = Decimal.Parse(no5G);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/Attenuator/Band_5G/GPIB_Number_5G " + ex);
            }

            /* Read 2.4G GPIB ip address */
            XmlNode nodeGpibBIP_5G = doc.SelectSingleNode("/CybertanATE/Attenuator/Band_5G/GPIB_IP_5G");
            /* Clear all items. Added by James */
            lboxRvRAtteuationSettingGPIBIP5G.Items.Clear() ;
            try
            {
                for (i = 1; i < nudRvRAtteuatorNumber5G.Value + 1; i++)
                {
                    string temp = "IP_" + i.ToString();
                    string ip = nodeGpibBIP_5G.SelectSingleNode("IP_" + i.ToString()).InnerText;
                    Debug.WriteLine("IP_" + i.ToString() + " :" + ip);
                    lboxRvRAtteuationSettingGPIBIP5G.Items.Add(ip);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/Attenuator/Band_5G/GPIB_IP_5G " + ex);
            }            

            /* Read Function Test Setting */
            /* Read Model */
            XmlNode nodeModel = doc.SelectSingleNode("/CybertanATE/FunctionTest/Model");
            try
            {
                string Name = nodeModel.SelectSingleNode("Name").InnerText;
                string SN = nodeModel.SelectSingleNode("SN").InnerText;
                string SWVer = nodeModel.SelectSingleNode("SWVer").InnerText;
                string HWVer = nodeModel.SelectSingleNode("HWVer").InnerText;

                cbox_Model_Name.SelectedIndex = Int32.Parse(Name);
                Debug.WriteLine("Name: " + Name + ": " + cbox_Model_Name.SelectedItem.ToString());
                Debug.WriteLine("SN: " + SN);
                Debug.WriteLine("SWVer: " + SWVer);
                Debug.WriteLine("HWVer: " + HWVer);    
                
                txt_Model_SerialNumber.Text = SN;
                txt_Model_SWVersion.Text = SWVer;
                txt_Model_HWVersion.Text = HWVer;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/FunctionTest/Model " + ex);
            }

            /* Read Router Setting */
            XmlNode nodeRouter = doc.SelectSingleNode("/CybertanATE/FunctionTest/Router");
            try
            {
                string Username = nodeRouter.SelectSingleNode("Username").InnerText;
                string Password = nodeRouter.SelectSingleNode("Password").InnerText;
                string DutIP = nodeRouter.SelectSingleNode("IP").InnerText;

                Debug.WriteLine("Username: " + Username);
                Debug.WriteLine("Password: " + Password) ;
                Debug.WriteLine("DUT IP: " + DutIP);

                txt_RouterSetting_UserName.Text = Username;
                txt_RouterSetting_Password.Text = Password;
                txt_RouterSetting_DUTIPAddress.Text = DutIP;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/FunctionTest/Router " + ex);
            }

            /* Read Client Setting */
            XmlNode nodeClient = doc.SelectSingleNode("/CybertanATE/FunctionTest/Client");
            try
            {
                string IP_24G = nodeClient.SelectSingleNode("IPaddr_2_4G").InnerText;
                string IP_5G = nodeClient.SelectSingleNode("IPaddr_5G").InnerText;

                Debug.WriteLine("IP_2_4G: "+IP_24G);
                Debug.WriteLine("IP_5G: "+IP_5G);

                txt_ClientSetting_24gIPaddress.Text = IP_24G;
                txt_ClientSetting_5gIPaddress.Text = IP_5G;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/FunctionTest/Client " + ex);
            }
        
            /* Read Chariot Folder */
            XmlNode nodeChariot = doc.SelectSingleNode("/CybertanATE/FunctionTest/Chariot");
            try
            {
                string cFile = nodeChariot.SelectSingleNode("File").InnerText;

                Debug.WriteLine("File: "+ cFile);

                txt_ChariotSetting_ChariotFolder.Text = cFile;

            }
            catch (Exception ex)
            {
                Debug.WriteLine("/CybertanATE/FunctionTest/Chariot " + ex);
            }
            
            return true;
        }

        private bool ConfigureRouter(string modelName,string band, string mode, int channel, string ssid, string security, string key)
        {
            string cgi= string.Empty;
            //string CgiHeader = "http://" + txt_RouterSetting_DUTIPAddress.Text + "/mfgtst.cgi?"; 
           
            switch (modelName)
            {


                case "TCH-2.4G":
                    string _ssidSecurity = "http://" + txtRvRTurnFunctionTestDUTIPAddress.Text + "";
                    string _channel = "http://" + txtRvRTurnFunctionTestDUTIPAddress.Text + "";
                    break;
                case "E8350":
                        cgi = "http://" + txt_RouterSetting_DUTIPAddress.Text + "/mfgtst.cgi?" + "sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=1";
                        if(band == "2.4G")
                        {
                            string iMode = mode.Substring(0, mode.Length - 1);
                            if (iMode == "40")
                            { /* when mode is 40M, need to switch channel to channel -2 */
                                if ((channel - 2) >= 0)
                                    cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + (channel - 2).ToString());
                                else
                                    cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=-1");
                            }
                            if (iMode == "20")
                            {
                                cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + channel.ToString());
                            }

                            cgi = cgi.Replace("sys_BandWidth=20", "sys_BandWidth=" + iMode);
                            cgi = cgi.Replace("sys_wlSSID=ssid", "sys_wlSSID=" + ssid);
                        }

                        /* cgi => sys_wlInterface=0&sys_wlMode=an-mixed&sys_BandWidth=20&sys_wlSSID=HWSQA5G&sys_wlChannel=1 */
                        if (band == "5G")
                        {
                            string[] s = mode.Split('_');
                            if (s[0] == "11N") cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=an-mixed");
                            if (s[0] == "11AC") cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=ac-mixed");
                            string iMode = s[1].Substring(0, s[1].Length - 1);

                            cgi = cgi.Replace("sys_wlInterface=0", "sys_wlInterface=1");
                            cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=an-mixed");
                            cgi = cgi.Replace("sys_BandWidth=20", "sys_BandWidth="+iMode );
                            cgi = cgi.Replace("sys_wlSSID=ssid", "sys_wlSSID=" + ssid);
                            cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + channel.ToString());
                        }

                        if (!SendCGICommand(cgi, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text, "200 OK"))
                        {
                            return false;
                        }
                        
                        Thread.Sleep(5000);
                    
                    break;
                case "EA8500X5":
                    string host = txt_RouterSetting_DUTIPAddress.Text;

                    /* configure router with telnet connection */
                    TelnetConnection tc = new TelnetConnection(host, 23);
                    string str = tc.Login("", "", 100);
                    Invoke(new SetTextCallBack(SetText), new object[] { str });
                    string prompt = str.TrimEnd();

                    //while connected
                    if (tc.IsConnected)
                    {
                        /* display server output */
                        str = tc.Read();
                        Invoke(new SetTextCallBack(SetText), new object[] { str });

                        /* send client input to server */
                        prompt = Console.ReadLine();
                        tc.WriteLine(prompt);

                        /* display server outpur */
                        str = tc.Read();
                        Invoke(new SetTextCallBack(SetText), new object[] { str });

                        //tc.WriteLine("ls");
                        //str = tc.Read();
                        //Invoke(new SetTextCallBack(SetText), new object[] { str });

                        /* Configure router band, channel and ssid */

                        str = " wifi_test.sh ap ";
                        if (band == "2.4G")
                        {
                            str += "2g ";
                            str += channel.ToString() + " " + ssid;
                            tc.WriteLine(str);
                            while (true)
                            {

                                Thread.Sleep(3000);
                                str = tc.Read();
                                if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
                                if (str.IndexOf("#") != -1) break;
                            }


                            /* Configure mode */
                            string smode = mode.Substring(0, 2);
                            if (smode != "40")
                            {
                                //str = "uci set wireless.@wifi-device[0].hwmode=11ng";
                                str = "uci set wireless.@wifi-iface[0].hwmode=11ng";
                                tc.WriteLine(str);
                                Thread.Sleep(1000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });
                                //str = "uci set wireless.@wifi-iface[0].htmode=VHT" + smode;
                                str = "uci set wireless.@wifi-device[0].htmode=HT" + smode;
                                tc.WriteLine(str);
                                Thread.Sleep(1000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });

                                /* Configure security */

                                /* Reset wifi */
                                str = "wifi down";
                                tc.WriteLine(str);
                                Thread.Sleep(3000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });
                                str = "wifi up";
                                tc.WriteLine(str);
                                while (true)
                                {

                                    Thread.Sleep(3000);
                                    str = tc.Read();
                                    if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
                                    if (str.IndexOf("#") != -1) break;
                                }
                                //Invoke(new SetTextCallBack(SetText), new object[] { str });


                                str = "iwpriv wifi0 setCountryID 843";
                                tc.WriteLine(str);
                                Thread.Sleep(3000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });



                            }

                            if (smode == "40")
                            {
                                //str = "uci set wireless.@wifi-device[0].hwmode=11ng";
                                str = "uci set wireless.@wifi-iface[0].hwmode=11ng";
                                tc.WriteLine(str);
                                Thread.Sleep(1000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });
                                str = "uci set wireless.@wifi-iface[0].htmode=VHT" + smode;
                                //str = "uci set wireless.@wifi-device[0].htmode=HT" + smode;
                                tc.WriteLine(str);
                                Thread.Sleep(1000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });

                                /* Configure security */

                                /* Reset wifi */
                                str = "wifi down";
                                tc.WriteLine(str);
                                Thread.Sleep(3000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });
                                str = "wifi up";
                                tc.WriteLine(str);
                                while (true)
                                {

                                    Thread.Sleep(3000);
                                    str = tc.Read();
                                    if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
                                    if (str.IndexOf("#") != -1) break;
                                }
                                //Invoke(new SetTextCallBack(SetText), new object[] { str });


                                str = "iwpriv wifi0 setCountryID 843";
                                tc.WriteLine(str);
                                Thread.Sleep(3000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });

                            }


                            /* Show wireless status */
                            str = "uci show wireless";
                            tc.WriteLine(str);
                            Thread.Sleep(3000);
                            str = tc.Read();
                            Invoke(new SetTextCallBack(SetText), new object[] { str });

                            /* Show speed */
                            str = "iwconfig";
                            tc.WriteLine(str);
                            Thread.Sleep(3000);
                            str = tc.Read();
                            Invoke(new SetTextCallBack(SetText), new object[] { str });
                        }

                        if (band == "5G")
                        {
                            str += "5g ";
                            str += channel.ToString() + " " + ssid;
                            tc.WriteLine(str);
                            while (true)
                            {

                                Thread.Sleep(3000);
                                str = tc.Read();
                                if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
                                if (str.IndexOf("#") != -1) break;
                            }



                            /* Configure mode */
                            string[] s = mode.Split('_');
                            string smode = s[1].Substring(0, 2);

                            if (smode == "20")
                            {
                                if (s[0] == "11N")
                                {
                                    str = "uci set wireless.@wifi-iface[1].hwmode=11an";
                                }
                                if (s[1] == "11AC")
                                {
                                    str = "uci set wireless.@wifi-iface[1].hwmode=11ac";
                                }

                                tc.WriteLine(str);
                                Thread.Sleep(1000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });


                                //str = "uci set wireless.@wifi-iface[1].htmode=VHT" + smode;
                                str = "uci set wireless.@wifi-device[1].htmode=HT" + smode;
                                tc.WriteLine(str);
                                Thread.Sleep(1000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });


                                /* Configure security */

                                /* Reset wifi */
                                str = "wifi down";
                                tc.WriteLine(str);
                                Thread.Sleep(3000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });

                                str = "wifi up";
                                tc.WriteLine(str);
                                while (true)
                                {

                                    Thread.Sleep(3000);
                                    str = tc.Read();
                                    if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
                                    if (str.IndexOf("#") != -1) break;
                                }

                                str = "iwpriv wifi0 setCountryID 843";
                                tc.WriteLine(str);
                                Thread.Sleep(3000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });

                            }

                            if (smode == "40")
                            {
                                if (s[0] == "11N")
                                {
                                    str = "uci set wireless.@wifi-iface[1].hwmode=11an";
                                }
                                if (s[1] == "11AC")
                                {
                                    str = "uci set wireless.@wifi-iface[1].hwmode=11ac";
                                }

                                tc.WriteLine(str);
                                Thread.Sleep(1000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });


                                //str = "uci set wireless.@wifi-iface[1].htmode=VHT" + smode;
                                str = "uci set wireless.@wifi-iface[1].htmode=VHT" + smode;
                                tc.WriteLine(str);
                                Thread.Sleep(1000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });


                                /* Configure security */

                                /* Reset wifi */
                                str = "wifi down";
                                tc.WriteLine(str);
                                Thread.Sleep(3000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });

                                str = "wifi up";
                                tc.WriteLine(str);
                                while (true)
                                {

                                    Thread.Sleep(3000);
                                    str = tc.Read();
                                    if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
                                    if (str.IndexOf("#") != -1) break;
                                }

                                str = "iwpriv wifi0 setCountryID 843";
                                tc.WriteLine(str);
                                Thread.Sleep(3000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });

                            }




                            if (smode == "80")
                            {
                                if (s[0] == "11N")
                                {
                                    str = "uci set wireless.@wifi-iface[1].hwmode=11an";
                                }
                                if (s[1] == "11AC")
                                {
                                    str = "uci set wireless.@wifi-iface[1].hwmode=11ac";
                                }

                                tc.WriteLine(str);
                                Thread.Sleep(1000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });


                                //str = "uci set wireless.@wifi-iface[1].htmode=VHT" + smode;
                                str = "uci set wireless.@wifi-iface[1].htmode=VHT" + smode;
                                tc.WriteLine(str);
                                Thread.Sleep(1000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });


                                /* Configure security */

                                /* Reset wifi */
                                str = "wifi down";
                                tc.WriteLine(str);
                                Thread.Sleep(3000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });

                                str = "wifi up";
                                tc.WriteLine(str);
                                while (true)
                                {

                                    Thread.Sleep(3000);
                                    str = tc.Read();
                                    if (str != "") Invoke(new SetTextCallBack(SetText), new object[] { str });
                                    if (str.IndexOf("#") != -1) break;
                                }

                                str = "iwpriv wifi0 setCountryID 843";
                                tc.WriteLine(str);
                                Thread.Sleep(3000);
                                str = tc.Read();
                                Invoke(new SetTextCallBack(SetText), new object[] { str });

                            }





                            /* Show wireless status */
                            str = "uci show wireless";
                            tc.WriteLine(str);
                            Thread.Sleep(3000);
                            str = tc.Read();
                            Invoke(new SetTextCallBack(SetText), new object[] { str });

                            /* Show speed */
                            str = "iwconfig";
                            tc.WriteLine(str);
                            Thread.Sleep(3000);
                            str = tc.Read();
                            Invoke(new SetTextCallBack(SetText), new object[] { str });


                        }

                    }

            // while connected   
                    //if (tc.IsConnected ) //&& prompt.Trim() != "exit")


                // send client input to server   
                    //  prompt = Console.ReadLine();
                    //  tc.WriteLine(prompt);

                // display server output   

            //    tc.WriteLine("wlanconfig ath0 list sta");
                    //    s = tc.Read();
                    //    textBox1.AppendText(s);

            //    tc.WriteLine("exit");



            //}

            //textBox1.AppendText("***DISCONNECTED");






                    /* CGI command has issue, replace with telnet */
                    // Shutdown interface 0 and 1
                    //cgi = "http://192.168.1.1:81/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=disabled";
                    //if (!SendCGICommand(cgi, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text, "wireless ConfigurationPass"))
                    //{
                    //    return false;
                    //}

                    //cgi = "http://192.168.1.1:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=disabled";
                    //if (!SendCGICommand(cgi, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text, "wireless ConfigurationPass"))
                    //{
                    //    return false;
                    //}

                    //cgi = "http://" + txt_RouterSetting_DUTIPAddress.Text + ":81/mfgtst.cgi?" + "sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=1";                    
                    //if (band == "2.4G")
                    //{
                    //    string iMode = mode.Substring(0, mode.Length - 1);
                    //    if (iMode == "40")
                    //    { /* when mode is 40M, need to switch channel to channel -2 */
                    //        if ((channel - 2) >= 0)
                    //            cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + (channel - 2).ToString());
                    //        else
                    //            cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=-1");
                    //    }
                    //    if (iMode == "20")
                    //    {
                    //        cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + channel.ToString());
                    //    }

                    //    cgi = cgi.Replace("sys_BandWidth=20", "sys_BandWidth=" + iMode);
                    //    cgi = cgi.Replace("sys_wlSSID=ssid", "sys_wlSSID=" + ssid);
                    //}

                    ///* cgi => sys_wlInterface=0&sys_wlMode=an-mixed&sys_BandWidth=20&sys_wlSSID=HWSQA5G&sys_wlChannel=1 */
                    //if (band == "5G")
                    //{
                    //    string[] s = mode.Split('_');
                    //    if (s[0] == "11N") cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=an-mixed");
                    //    if (s[0] == "11AC") cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=ac-mixed");
                    //    string iMode = s[1].Substring(0, s[1].Length - 1);

                    //    cgi = cgi.Replace("sys_wlInterface=0", "sys_wlInterface=1");
                    //    cgi = cgi.Replace("sys_wlMode=11n-2G-only", "sys_wlMode=an-mixed");
                    //    cgi = cgi.Replace("sys_BandWidth=20", "sys_BandWidth=" + iMode);
                    //    cgi = cgi.Replace("sys_wlSSID=ssid", "sys_wlSSID=" + ssid);
                    //    cgi = cgi.Replace("sys_wlChannel=1", "sys_wlChannel=" + channel.ToString());
                    //}

                    //if (!SendCGICommand(cgi, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text, "wireless ConfigurationPass"))
                    //{
                    //    return false;
                    //}
                    else return false;
                    break;



                case "EA8500X8":
                    string hostX8 = EA8500X8.host;
                    hostX8 = hostX8.Replace("DUTIPaddr", txt_RouterSetting_DUTIPAddress.Text);
                    string data = string.Empty;
                    hostX8.Replace("DUTIPaddr", txt_RouterSetting_DUTIPAddress.Text);
                    if (band == "2.4G")
                    {
                        string iMode = mode.Substring(0, mode.Length - 1);
                        

                        if (iMode == "40")
                        {
                            data = EA8500X8.data_2_4G_11n_40M;
                        }
                        if (iMode == "20")
                        {
                            data = EA8500X8.data_2_4G_11n_20M;
                        }

                        data = data.Replace("SSIDName", ssid);
                        data = data.Replace("channel_3", channel.ToString());
                        if(security != "None")
                        {
                            data = data.Replace(@"None""}}]}}]", security+@""",""wpaPersonalSettings"":{""passphrase"":"""+passphrase+@"""}}}]}}]");
                        }

                        if (!WebHttpPostEA8500X8(hostX8, data, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text))
                        {
                            return false;
                        }
                        //WebHttpPostEA8500X8(hostX8, data, "admin", "admin");


                    }

                    if (band == "5G")
                    {
                        switch (mode)
                        {  /* "11N_20M", "11N_40M", "11AC_20M", "11AC_40M", "11AC_80M" */
                            case "11N_20M":
                                data = EA8500X8.data_5G_11n_20M;
                            break ;
                            
                            case "11N_40M":
                                data = EA8500X8.data_5G_11n_40M;
                            break;
                            
                            case "11AC_20M":
                                data = EA8500X8.data_5G_11ac_20M;
                            break;
                            
                            case "11AC_40M":
                                data = EA8500X8.data_5G_11ac_40M;
                            break;
                            
                            case "11AC_80M":
                                data = EA8500X8.data_5G_11ac_80M;
                            break;

                            default:
                                return false;
                            break;
                        }


                        data = data.Replace("SSIDName", ssid);
                        data = data.Replace("channel_3", channel.ToString());
                        if (security != "None")
                        {
                            data = data.Replace(@"None""}}]}}]", security + @""",""wpaPersonalSettings"":{""passphrase"":""" + passphrase + @"""}}}]}}]");
                        }

                        if (!WebHttpPostEA8500X8(hostX8, data, txt_RouterSetting_UserName.Text, txt_RouterSetting_Password.Text))
                        {
                            return false;
                        }
                    }



                    break;
                
                default:
                    return false;           
            }

            return true;        
        }

        private bool ReadGpibIP(string gpib_interface, string gpib_ip)
        {            
            int no_2_4G = (int)nudRvRAtteuatorNumber24G.Value;
            int no_5G = (int)nudRvRAtteuatorNumber5G.Value;

            if (no_2_4G != 0)
            {
                int i = 0;
                a11713A_2_4G = new Agilent11713A[no_2_4G];
                foreach (string s in lboxRvRAtteuationSettingGPIBIP24G.Items)
                {
                    string resourceName = "GPIB" + nudRvRGPIBInterface24G.Value.ToString() + "::" + s + "::INSTR";
                    a11713A_2_4G[i] = new Agilent11713A(resourceName);
                    if (a11713A_2_4G[i] == null) return false;
                }
            }

            if (no_5G != 0)
            {
                int i = 0;
                a11713A_5G = new Agilent11713A[no_5G];
                foreach (string s in lboxRvRAtteuationSettingGPIBIP5G.Items)
                {
                    string resourceName = "GPIB" + nudRvRGPIBInterface5G.Value.ToString() + "::" + s + "::INSTR";
                    a11713A_5G[i] = new Agilent11713A(resourceName);
                    if (a11713A_5G[i] == null) return false;
                    i++;
                }
            }
            return true;
        }

        private bool OpenGPIBport()
        {
            int no_2_4G = (int)nudRvRAtteuatorNumber24G.Value;
            int no_5G = (int)nudRvRAtteuatorNumber5G.Value;

            if (no_2_4G != 0)
            {
                int i = 0;
                a11713A_2_4G = new Agilent11713A[no_2_4G];
                foreach (string s in lboxRvRAtteuationSettingGPIBIP24G.Items)
                {
                    //string resourceName = "GPIB" + nudRvRGPIBInterface24G.Value.ToString() + "::" + s + "::INSTR";
                    string resourceName = "GPIB" + nudRvRGPIBInterface24G.Value.ToString() + "::" + s;
                    a11713A_2_4G[i] = new Agilent11713A(resourceName);
                    if (a11713A_2_4G[i] == null) return false;
                    i++;
                }
            }

            if (no_5G != 0)
            {
                int i = 0;
                a11713A_5G = new Agilent11713A[no_5G];
                foreach (string s in lboxRvRAtteuationSettingGPIBIP5G.Items)
                {
                    string resourceName = "GPIB" + nudRvRGPIBInterface5G.Value.ToString() + "::" + s + "::INSTR";
                    a11713A_5G[i] = new Agilent11713A(resourceName);
                    if (a11713A_5G[i] == null) return false;
                    i++;
                }
            }

            return true;
        }

        private void ReadAtteunatorValue()
        {
            Attenuation_buttonValue_2_4G = new decimal[]{nudRvRAtteuation24GX1.Value,
                nudRvRAtteuation24GX2.Value,
                nudRvRAtteuation24GX3.Value,
                nudRvRAtteuation24GX4.Value,
                nudRvRAtteuation24GX5.Value,
                nudRvRAtteuation24GX6.Value,
                nudRvRAtteuation24GX7.Value,
                nudRvRAtteuation24GX8.Value};

            Attenuation_buttonValue_5G = new decimal[]{nudRvRAtteuation5GX1.Value ,
                nudRvRAtteuation5GX2.Value ,
                nudRvRAtteuation5GX3.Value ,
                nudRvRAtteuation5GX4.Value ,
                nudRvRAtteuation5GX5.Value ,
                nudRvRAtteuation5GX6.Value ,
                nudRvRAtteuation5GX7.Value ,
                nudRvRAtteuation5GX8.Value };       
        }

        private bool CheckAtteuatorOn()
        {
            int nodevice = 0;
            int nosetting;
            nosetting = lboxRvRAtteuationSettingGPIBIP24G.Items.Count + lboxRvRAtteuationSettingGPIBIP5G.Items.Count;
            string[] resource = FindGPIBResource();
           
            foreach (string s in lboxRvRAtteuationSettingGPIBIP24G.Items)
            {                
                foreach (string r in resource)
                {
                    // Needs to modify it because this is not a good method
                    char[] delimiterChars = { ':'};
                    string[] words = r.Split(delimiterChars);

                    //if (s == r)
                    if (s == words[2])                  
                        nodevice++;
                }
            }

            foreach (string s in lboxRvRAtteuationSettingGPIBIP5G.Items)
            {                
                foreach (string r in resource)
                {
                    //if (s == r) nodevice++;
                    // Needs to modify it because this is not a good method
                    char[] delimiterChars = { ':' };
                    string[] words = r.Split(delimiterChars);

                    //if (s == r)
                    if (s == words[2])
                        nodevice++;
                }
            }

            if(nosetting == nodevice) return true;
            else return false ;
        }

        private void ToggleFunctionTestRvR()
        {
            ToggleRvRFunctionTestControl(true);
            Debug.WriteLine("Toggle");
        }

        private string GetWiFiData()
        {
            string returnData = "";            
            try
            {
                UdpClient receivingUdpClient = new UdpClient(9051);
                IPEndPoint RemoteIpEndPoint = new IPEndPoint(IPAddress.Any, 0);

                Byte[] receiveBytes = receivingUdpClient.Receive(ref RemoteIpEndPoint);

                returnData = Encoding.ASCII.GetString(receiveBytes);

                Debug.WriteLine("This is the message you received " + returnData.ToString());
                Debug.WriteLine("This message was sent from " + RemoteIpEndPoint.Address.ToString() +
                    " on their point number" +
                    RemoteIpEndPoint.Port.ToString());

                receivingUdpClient.Close();
                receivingUdpClient = null;
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.ToString());
            }

            return returnData;
        }

        private void WriteCSV(string[] files, string outputfile)
        {
            string value;
            StringBuilder csv = new StringBuilder();
            string newline;

            if (File.Exists(outputfile)) 
                File.Delete(outputfile);

            int i = 0;
            //for (int i = 0; i < files.Length; i++)
            foreach(string s in files)
            {
                //value = csv_GetSpecificValue(files[i], 15, 16);

                 string[] Read_y_point = File.ReadAllLines(s);
                 int y_length = Read_y_point.Length;
                 string[] Read_y_point_value = Read_y_point[y_length - 1].Split(',');
                 double Y_value = Convert.ToDouble(Read_y_point_value[9]);
                 value = Y_value.ToString();


                newline = string.Format("{0},{1}{2}", Path.GetFileName(files[i++]), value, Environment.NewLine);
                csv.Append(newline);
            }

            //before your loop
            /*
            var csv = new StringBuilder();

            //in your loop
            var first = reader[0].ToString();
            var second = image.ToString();
            var newLine = string.Format("{0},{1}{2}", first, second, Environment.NewLine);
            csv.Append(newLine);

            */
            //after your loop
            File.WriteAllText(outputfile, csv.ToString());
        }

        private void WriteResultCSV(string files, string outputfile)
        {
            //StringBuilder csv = new StringBuilder();
            //string newline;
            /*
            if (File.Exists(outputfile))
                File.Open(outputfile, FileMode.Open);
            */
            //newline = string.Format("{0},{1}{2}", files, WiFiData, Environment.NewLine);
            //linkrate_csv.Append(newline);

            //File.WriteAllText(outputfile, linkrate_csv.ToString());
            /*
            if (!File.Exists(outputfile))
            {
                FileStream filestream = File.Create(outputfile);
                filestream.Close();
            }
             * */
        }

        private string csv_GetSpecificValue(string path, int row, int col)
        {
            string val = string.Empty;

            StreamReader file = new StreamReader(path);
            int count = 1;
            string line = string.Empty;
            while (!file.EndOfStream)
            {
                if (count == row)
                {
                    line = file.ReadLine();
                    val = line.Split(',')[col];
                    file.Close();
                    break;
                }
                file.ReadLine();
                count++;
            }
            return val;
        }

        private bool WebHttpPostEA8500X8(string head, string data, string username, string password)
        {
            byte[] bs = Encoding.ASCII.GetBytes(data);

            //HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create("http://www.google.com/intl/zh-CN/");
            //HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create("http://192.168.1.1");
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(head);
            req.Method = "POST";
            req.ContentType = "application/json; charset=UTF-8";

            //req.ContentType = "text/xml";
            //req.ContentType = "application/x-www-form-urlencoded";


            req.ContentLength = bs.Length;
            //req.Headers.Add("
            req.Headers.Add("x-jnap-authorization", "Basic YWRtaW46YWRtaW4=\r\n");
            req.Headers.Add("x-jnap-action", "http://linksys.com/jnap/core/Transaction\r\n");
            req.Headers.Add("x-requested-with", "XMLHttpRequest\r\n");

            req.Headers.Add("Authorization", "Basic YWRtaW46YWRtaW4=\r\n");
            //req.Headers.Add("Credentials", "admin:admin");
            req.Headers.Add("Credentials", username + ":" + password);

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
                    //textBox1.AppendText(strHtml);

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

        private string ConfigureWithTelnet(string host, string username, string password)
        {
            string str = string.Empty;
            return str;
        }

        private void chkbox_RouterSetting_ConfigRouterManually_CheckStateChanged(object sender, EventArgs e)
        {
            if (chkbox_RouterSetting_ConfigRouterManually.Checked)
            {
                txt_Temp_ModelName.Visible = true;
                lab_Temp_ModelName.Visible = true;
            }
            else
            {
                txt_Temp_ModelName.Visible = false;
                lab_Temp_ModelName.Visible = false;
            }
        }

        private void RvRTabHideAllPages()
        {
            tp_RvRTestCondition.Parent = null; //Hide tabPage
            tp_RvRAttenuatorSetting.Parent = null;
            tp_RvRDeviceTest.Parent = null;
            tp_RvRTurnFunctionTest.Parent = null;
            tp_RvRFunctionTest.Parent = null;
            tp_RvRTestResultInfo.Parent = null;
            tp_RvRTurnTestCondition.Parent = null;
            tp_RvRTestResultChart.Parent = null;
        }

        private void ForTest()
        {
            string[] a= new string[5] ;
            string[,] b= new string[4,5] ;

            int length = a.Length;
            int l = b.Length;
            int le = b.Rank;

            int ll = b.GetLength(0);
            int lf = b.GetLength(1);
            for (int dim = 1; dim <= le; dim++)
            {
                //string k = 
            }

            //int d = 0;
        }
    } //End of class 
}
