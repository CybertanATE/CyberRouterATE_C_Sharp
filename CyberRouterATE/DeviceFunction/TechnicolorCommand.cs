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
using System.Net.NetworkInformation;
using System.IO;
using AgilentInstruments;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        int[] TCH5G20MChannel =new int[] {36, 40, 44, 48, 52, 56, 60, 64, 100, 104, 108, 112, 116, 120, 124, 128, 132, 136, 140, 149, 153, 157, 161, 165} ;
        int[] TCH5G40MLowerChannel =new int[] {36, 44, 52, 60, 100, 108, 116, 124, 132, 149, 157} ;
        int[] TCH5G40MUpperChannel = new int[] {40, 48, 56, 64, 104, 112, 120, 128, 136, 153, 161 };

        public bool TchLogin(string host, string userName, string passWord)
        {
            return false;
        }    

        public bool Tch24GDeviceConfigure(string hostIP, string userName, string passWord, WifiBasic wifi)
        {
            if (wifi.band == "5G") return false;
            //if (wifi.mode.IndexOf("40M") >= 0) return false; //無2.4G 40M 選項可設

            string response = string.Empty;
            
            /* Login in process */
            string _uri = @"http://" + hostIP;            
            string _agent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; InfoPath.3)";
            string _contentType = "application/x-www-form-urlencoded";
            bool _keepAlive = true;
            bool _Auth = false;
            string _username = "";
            string _password = "";
            string _Type = "";
            string _CSRFValue = string.Empty;           

            /* Get CSRFVaule */ 
            WebHttpGetData(_uri, _agent, _keepAlive, _Auth, _username, _password, "", ref response);
            

            int index = response.ToLower().IndexOf("csrfvalue");
            if (index == -1)
            {
                Debug.WriteLine("CSRFValue not found");
                TCHLogout(hostIP);
                return false;
            }

            string _csrfValuetmp = response.Substring(index + 9, 30);
            index = _csrfValuetmp.ToLower().IndexOf("value=");

            if (index == -1)
            {
                Debug.WriteLine("Value not found");
                TCHLogout(hostIP);
                return false;
            }

            _csrfValuetmp = _csrfValuetmp.Substring(index + 6, 10);

            string[] str = _csrfValuetmp.Split('>');

            _CSRFValue = str[0];

            _uri = @"http://"+hostIP + @"/goform/login";
            string _data = "CSRFValue=" + _CSRFValue + "&loginUsername=" + userName +"&loginPassword=" + passWord + "&logoffUser=0";
            _agent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; InfoPath.3)";
            _contentType = "application/x-www-form-urlencoded";

            _keepAlive = true;
            _Auth = false;
            _username = "";
            _password = "";
            _Type = "";
            try
            {
                WebHttpPostData(_uri, _data, _agent, _contentType, _keepAlive, _Type, _Auth, _username, _password, ref response);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Login Fail: " + ex.ToString());
                return false;
            }

            if (response.ToLower().IndexOf("wireless") < 0)
            {
                TCHLogout(hostIP);
                return false;
            }
            if (response.ToLower().IndexOf("firewall") < 0)
            {
                TCHLogout(hostIP);
                return false;
            }

            /*Configure Channel and Band*/
            _uri = @"http://" + hostIP + @"/goform/wlanRadio";
            //_uri = "http://192.168.0.1/goform/wlanRadio";

            string _outputPower = "OutputPower=100";
            string _band = "Band=2"; //
            string _sideBand = "NSideband=-1";
            string _bandWidth = "NBandwidth=20";
            string _channelNumber = "ChannelNumber=10";
            string _coexitence = "ObssCoexistence=1";
            string _stbctx = "STBCTx=0";
            string _restoredWirelessDefaults = "restoreWirelessDefaults=0";
            string _commitwlanRadio = "commitwlanRadio=0";
            string _scanActions = "scanActions = 0";
            string _connectString = "&";
            string _nMode = string.Empty;


            if (wifi.mode.IndexOf("20M") >= 0)
            {
                _bandWidth = "NBandwidth=20";
                _sideBand = "NSideband=-1";

            }
            if (wifi.mode.IndexOf("40M") >= 0)
            {
                _bandWidth = "NBandwidth=40";
                if (wifi.channel > 9)
                {
                    _sideBand = "NSideband=1";
                }
                else
                {
                    _sideBand = "NSideband=-1";
                }
            }
            //if (wifi.mode.IndexOf("80M") >= 0) _bandWidth = "NBandwidth=80";

            if (wifi.mode.IndexOf("11N") >= 0) _nMode = "NMode=1";
            else
            {
                if (wifi.mode.IndexOf("11AC") >= 0) _nMode = "NMode=2";
                else _nMode = "NMode=0";
            }


            /* channel 5~13  -side upper, lower-1~9 */
            string[] _wlanRadioData = new string[]
            {
                "OutputPower=100", //Output Power : 100%
                "Band=2",   //Band : 2-2.4G
                _nMode,
                //"NMode=1",  //NMode : 0-OFF, 1-Auto
                _bandWidth,
                //"NBandwidth=20",  //NBandWidth : 20-20MHz, 40-40MHz                
                "ChannelNumber=" + wifi.channel.ToString(), //Channel number
                "ObssCoexistence=1", //Co-existence : 0-Disable, 1 -Enable, Use Default
                "STBCTx=0",   //STBC Tx: 0-Auto, 1-On, 2-Off , use default
                "restoreWirelessDefaults=0", //USe Default
                "commitwlanRadio=1", //USe Default
                "scanActions = 0" //USe Default           
            };

            _data = string.Empty;

            for (int i = 0; i < _wlanRadioData.Length - 1; i++)
            {
                _data += _wlanRadioData[i] + _connectString;
            }

            _data += _wlanRadioData[_wlanRadioData.Length - 1];            
            _contentType = "application/x-www-form-urlencoded";

            _keepAlive = true;
            WebHttpPostData(_uri, _data, _agent, _contentType, _keepAlive, _Type, _Auth, _username, _password, ref response);            
            Thread.Sleep(3000);

            if (response.ToLower().IndexOf("2.4 ghz") < 0)
            {
                TCHLogout(hostIP);
                return false;
            }
            //if (response.ToLower().IndexOf("firewall") < 0) return false;
            
            /* Configure SSID and security */
            _uri = @"http://" + hostIP + @"/goform/wlanPrimaryNetwork";

            //_agent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36";
            _contentType = "application/x-www-form-urlencoded";
            _keepAlive = true;

            /* Configure SSID and securtiyt */
            string _wpa2pskauth = "Wpa2PskAuth=0";
            string _wpa2enc = "WpaEncryption=0";
            string _wpa2key = "WpaPreSharedKey=" + wifi.passphrase;
            if (wifi.security.ToLower() == "none")
            {
                _wpa2pskauth = "Wpa2PskAuth=0";
                _wpa2enc = "WpaEncryption=0";
                _wpa2key = "WpaPreSharedKey=" + wifi.passphrase;
            }
            if (wifi.security.ToLower() == "wpa2-personal")
            {
                _wpa2pskauth = "Wpa2PskAuth=1";
                _wpa2enc = "WpaEncryption=0";
                _wpa2key = "WpaPreSharedKey=" + wifi.passphrase;
            }

            string[] _wlanPrimaryNetwork = new string[]
            {
                "PrimaryNetworkEnable=1",
                "ServiceSetIdentifier=" + wifi.ssid,
                "ClosedNetwork=0",
                "ApIsolate=0",
                "WpaAuth=0",
                "WpaPskAuth=0",
                "Wpa2Auth=0",
                _wpa2pskauth,
                _wpa2enc,
                _wpa2key,
                //"Wpa2PskAuth=1", //Wpa2PskAuth=0 =>none, Wpa2PskAuth=1=>WPA2PSK
                //"WpaEncryption=0",
                //"WpaPreSharedKey=1234567890",
                "ShowWpaKey=0x01",
                "RadiusServer=0.0.0.0",
                "RadiusPort=1812",
                "RadiusKey=",
                "WpaRekeyInterval=0",
                "WpaReauthInterval=3600",
                "SharedKeyAuthentication=0",
                "w802_1xAuthentication=0",
                "NetworkKey1=",
                "NetworkKey2=",
                "NetworkKey3=",
                "NetworkKey4=",
                "DefaultSecretKey=1",
                "WepPassPhrase=",
                "GenerateWepKeys=0",
                "commitwlanPrimaryNetwork=0"
            };

            _data = string.Empty;

            for (int i = 0; i < _wlanPrimaryNetwork.Length - 1; i++)
            {
                _data += _wlanPrimaryNetwork[i] + "&";
            }

            _data += _wlanPrimaryNetwork[_wlanPrimaryNetwork.Length - 1];

            WebHttpPostData(_uri, _data, _agent, _contentType, _keepAlive, _Type, _Auth, _username, _password, ref response);
            Thread.Sleep(3000);

            if (response.ToLower().IndexOf(wifi.ssid.ToLower()) < 0)
            {
                TCHLogout(hostIP);
                return false;
            }
            //if (response.ToLower().IndexOf(wifi.ssid.ToLower()) < 0) return false;            

            TCHLogout(hostIP);
            return true;
        }

        public bool Tch5GDeviceConfigure(string hostIP, string userName, string passWord, WifiBasic wifi)
        {
            string response = string.Empty;

            string _wirelessEnable = "WirelessEnable=1";
            string _outputPower = "OutputPower=100";
            string _band = "Band=2"; //
            string _sideBand = "NSideband=-1";
            string _bandWidth = "NBandwidth=20";
            string _channelNumber = "ChannelNumber=10";
            string _coexitence = "ObssCoexistence=1";
            string _stbctx = "STBCTx=0";
            string _restoredWirelessDefaults = "restoreWirelessDefaults=0";
            string _commitwlanRadio = "commitwlanRadio=1";
            string _scanActions = "scanActions = 0";
            string _connectString = "&";
            string _nMode = "NMode=1";
            string _regulatoryMode = "RegulatoryMode=0";            

            /* Configure SSID and securtiyt */
            string _wpa2pskauth = "Wpa2PskAuth=0";
            string _wpa2enc = "WpaEncryption=0";
            string _wpa2key = "WpaPreSharedKey=12345678";
            string _ssid = "ServiceSetIdentifier=" + wifi.ssid.Trim();
                        
            /* Login in process */
            string _uri = @"http://" + hostIP;
            string _agent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; InfoPath.3)";
            string _contentType = "application/x-www-form-urlencoded";
            bool _keepAlive = true;
            bool _Auth = false;
            string _username = "";
            string _password = "";
            string _Type = "";
            string _CSRFValue = string.Empty;

            /* Get CSRFVaule */
            WebHttpGetData(_uri, _agent, _keepAlive, _Auth, _username, _password, "", ref response);            

            int index = response.ToLower().IndexOf("csrfvalue");
            if (index == -1)
            {
                Debug.WriteLine("CSRFValue not found");
                return false;
            }

            string _csrfValuetmp = response.Substring(index + 9, 30);
            index = _csrfValuetmp.ToLower().IndexOf("value=");

            if (index == -1)
            {
                Debug.WriteLine("Value not found");
                return false;
            }

            _csrfValuetmp = _csrfValuetmp.Substring(index + 6, 10);

            string[] str = _csrfValuetmp.Split('>');

            _CSRFValue = str[0];

            _uri = @"http://" + hostIP + @"/goform/login";
            string _data = "CSRFValue=" + _CSRFValue + "&loginUsername=" + userName + "&loginPassword=" + passWord + "&logoffUser=0";
            _agent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; InfoPath.3)";
            _contentType = "application/x-www-form-urlencoded";
           
            try
            {
                WebHttpPostData(_uri, _data, _agent, _contentType, _keepAlive, _Type, _Auth, _username, _password, ref response);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Login Fail: " + ex.ToString());
                return false;
            }

            if (response.ToLower().IndexOf("wireless") < 0) return false;
            if (response.ToLower().IndexOf("firewall") < 0) return false;

            /*========= 2.4G Process =========*/
            if (wifi.band == "2.4G")
            {
                /*Configure 2.4G Channel and Band*/
                _uri = @"http://" + hostIP + @"/goform/wlanRadio";
                //_uri = "http://192.168.0.1/goform/wlanRadio";
                
                if (wifi.mode.IndexOf("20M") >= 0)
                {
                    _bandWidth = "NBandwidth=20";
                    _sideBand = "NSideband=-1";

                }
                /* 40M: channel 5~13  -side upper, lower-1~9 */
                if (wifi.mode.IndexOf("40M") >= 0)
                {
                    _bandWidth = "NBandwidth=40";
                    if (wifi.channel > 7)
                    {
                        _sideBand = "NSideband=1";
                    }
                    else
                    {
                        _sideBand = "NSideband=-1";
                    }
                }
                //if (wifi.mode.IndexOf("80M") >= 0) _bandWidth = "NBandwidth=80";

                if (wifi.mode.IndexOf("11N") >= 0) _nMode = "NMode=1";
                else
                {
                    if (wifi.mode.IndexOf("11AC") >= 0) _nMode = "NMode=2";
                    else _nMode = "NMode=0";
                }

                _channelNumber = "ChannelNumber=" + wifi.channel.ToString();

                string[] _wlanRadioData = new string[]
                {
                    _wirelessEnable,
                    _outputPower,
                    _band,
                    _nMode,     //"NMode=1",
                    _bandWidth,
                    _sideBand, /* Only for 40M (Value: 1, -1) value = -1 when 20M*/
                    _channelNumber,
                    _regulatoryMode,
                    _coexitence,
                    _stbctx,
                    _restoredWirelessDefaults,
                    _commitwlanRadio,
                    _scanActions            
                };


                _data = string.Empty;

                for (int i = 0; i < _wlanRadioData.Length - 1; i++)
                {
                    _data += _wlanRadioData[i] + _connectString;
                }

                _data += _wlanRadioData[_wlanRadioData.Length - 1];
                _contentType = "application/x-www-form-urlencoded";

                _keepAlive = true;
                WebHttpPostData(_uri, _data, _agent, _contentType, _keepAlive, _Type, _Auth, _username, _password, ref response);
                Thread.Sleep(3000);

                if (response.ToLower().IndexOf("2.4 ghz") < 0) return false;
                //if (response.ToLower().IndexOf("firewall") < 0) return false;

                /* Configure 2.4G SSID and securtiyt */
                _uri = @"http://" + hostIP + @"/goform/wlanPrimaryNetwork";

                if (wifi.security.ToLower() == "none")
                {
                    _wpa2pskauth = "Wpa2PskAuth=0";
                    _wpa2enc = "WpaEncryption=0";
                    _wpa2key = "WpaPreSharedKey=12345678";
                }
                if (wifi.security.ToLower() == "wpa2-personal")
                {
                    _wpa2pskauth = "Wpa2PskAuth=1";
                    _wpa2enc = "WpaEncryption=2";
                    _wpa2key = "WpaPreSharedKey=" + wifi.passphrase;
                }

                string[] _wlanPrimaryNetwork = new string[]
                {
                    "PrimaryNetworkEnable=1",
                    _ssid,
                    "ClosedNetwork=0",
                    "BssModeRequired=0",
                    "ApIsolate=0",                
                    "WpaPskAuth=0", 
                    _wpa2pskauth,
                    _wpa2enc,
                    _wpa2key,                              
                    "WpaRekeyInterval=0",                
                    "GenerateWepKeys=0",
                    "WepKeysGenerated=0",
                    "commitwlanPrimaryNetwork=0",
                    "AutoSecurity=1"
                };

                _data = string.Empty;

                for (int i = 0; i < _wlanPrimaryNetwork.Length - 1; i++)
                {
                    _data += _wlanPrimaryNetwork[i] + "&";
                }

                _data += _wlanPrimaryNetwork[_wlanPrimaryNetwork.Length - 1];
                //_data = _outputPower + _band;
                //_data = "OutputPower=100&Band=2&NMode=1&NBandwidth=20&ChannelNumber=10&ObssCoexistence=1&STBCTx=0&restoreWirelessDefaults=0&commitwlanRadio=0&scanActions=0";
                //_agent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36";
                _contentType = "application/x-www-form-urlencoded";

                _keepAlive = true;
                WebHttpPostData(_uri, _data, _agent, _contentType, _keepAlive, _Type, _Auth, _username, _password, ref response);
                Thread.Sleep(3000);

                if (response.ToLower().IndexOf(wifi.ssid.ToLower()) < 0) return false;
            }



            /*========= 5G Process =========*/
            if (wifi.band == "5G")
            {
                /*Configure 5G Channel and Band*/
                _uri = @"http://" + hostIP + @"/goform/wlan5gRadio";

                string _wireless5gEnable = "Wireless5gEnable=1";
                string _outputPower5g = "OutputPower5g=100";
                string _band5g = "Band5g=1"; //
                string _sideBand5g = "NSideband5g=-1";
                string _bandWidth5g = "NBandwidth5g=20";
                string _channelNumber5g = "ChannelNumber5g=157";
                string _coexitence5g = "ObssCoexistence5g=0";
                string _stbctx5g = "STBCTx5g=0";
                string _restoredWireless5gDefaults = "restoreWireless5gDefaults=0";
                string _commitwlan5gRadio = "commitwlan5gRadio=1";
                string _scanActions5g = "scanActions5g=0" ;
                string _nMode5g = "NMode5g=1";
                string _regulatoryMode5g = "RegulatoryMode5g=0";
  
                if (wifi.mode.IndexOf("20M") >= 0)
                {
                    _bandWidth5g = "NBandwidth5g=20";
                    _sideBand5g = "NSideband5g=-1";
                }
                /* 40M: channel 5~13  -side upper, lower-1~9 */
                if (wifi.mode.IndexOf("40M") >= 0)
                {
                    _bandWidth5g = "NBandwidth5g=40";
                    int mode = 0;
                    foreach (int channel in TCH5G40MUpperChannel)
                    {
                        if (wifi.channel == channel)
                        {
                            mode = 1;
                            _sideBand5g = "NSideband5g=1";
                        }
                    }
                    foreach (int channel in TCH5G40MLowerChannel)
                    {
                        if (wifi.channel == channel)
                        {
                            mode = 2;
                            _sideBand5g = "NSideband5g=-1";
                        }
                    }
                    if (mode == 0)
                    {
                        Debug.WriteLine("Channel is out of band");
                        return false;
                    }                    
                }
                //if (wifi.mode.IndexOf("80M") >= 0) _bandWidth = "NBandwidth=80";

                if (wifi.mode.IndexOf("11N") >= 0) _nMode5g = "NMode5g==1";
                else
                {
                    if (wifi.mode.IndexOf("11AC") >= 0) _nMode5g = "NMode5g==2";
                    else _nMode = "NMode5g==0";
                }

                _channelNumber5g = "ChannelNumber5g=" + wifi.channel.ToString();
                   
                //string[] _wlan5gRadio = new string[]
                //{
                //    "Wireless5gEnable=1",
                //    "OutputPower5g=100",
                //    "Band5g=1",
                //    "NMode5g=1",
                //    "NBandwidth5g=20",
                //    "NSideband5g=-1",/* Only for 40M, value = -1 when 20M*/
                //    "ChannelNumber5g=157",
                //    "RegulatoryMode5g=0",
                //    "ObssCoexistence5g=0",
                //    "STBCTx5g=0",
                //    "restoreWireless5gDefaults=0",
                //    "commitwlan5gRadio=1",
                //    "scanActions5g=0"            
                //};

                string[] _wlan5gRadioData = new string[]
                {
                    _wireless5gEnable,
                    _outputPower5g,
                    _band5g,
                    _nMode5g,
                    _bandWidth5g,
                    _sideBand5g,/* Only for 40M, value = -1 when 20M*/
                    _channelNumber5g,
                    _regulatoryMode5g,
                    _coexitence5g,
                    _stbctx5g,
                    _restoredWireless5gDefaults,
                    _commitwlan5gRadio,
                    _scanActions5g            
                };

                _data = string.Empty;

                for (int i = 0; i < _wlan5gRadioData.Length - 1; i++)
                {
                    _data += _wlan5gRadioData[i] + "&";
                }

                _data += _wlan5gRadioData[_wlan5gRadioData.Length - 1];
                //_data = _outputPower + _band;
                //_data = "OutputPower=100&Band=2&NMode=1&NBandwidth=20&ChannelNumber=10&ObssCoexistence=1&STBCTx=0&restoreWirelessDefaults=0&commitwlanRadio=0&scanActions=0";
                //_agent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36";
                _contentType = "application/x-www-form-urlencoded";

                _keepAlive = true;
                WebHttpPostData(_uri, _data, _agent, _contentType, _keepAlive, _Type, _Auth, _username, _password, ref response);
                Thread.Sleep(3000);

                if (response.ToLower().IndexOf("5 ghz") < 0) return false;
                

                /* Configure 5G SSID and securtiyt */
                _uri = @"http://" + hostIP + @"/goform/wlan5gPrimaryNetwork";                

                string _wpa2pskauth5g = "Wpa2PskAuth5g=0";
                string _wpa2enc5g = "WpaEncryption5g=0";
                string _wpa2key5g = "WpaPreSharedKey5g=12345678";
                string _ssid5g = "ServiceSetIdentifier5g=" + wifi.ssid.Trim();

                if (wifi.security.ToLower() == "none")
                {
                    _wpa2pskauth5g = "Wpa2PskAuth5g=0";
                    _wpa2enc5g = "WpaEncryption5g=0";
                    _wpa2key5g = "WpaPreSharedKey5g=12345678";
                }
                if (wifi.security.ToLower() == "wpa2-personal")
                {
                    _wpa2pskauth5g = "Wpa2PskAuth5g=1";
                    _wpa2enc5g = "WpaEncryption5g=2";
                    _wpa2key5g = "WpaPreSharedKey5g=" + wifi.passphrase;
                }
                                
                //string[] _wlan5gPrimaryNetworkData = new string[]
                //{
                //    "PrimaryNetworkEnable5g=1",
                //    "ServiceSetIdentifier5g=TECH_6G",
                //    "ClosedNetwork5g=0",
                //    "BssModeRequired5g=0",
                //    "ApIsolate5g=0",
                //    "WpaAuth5g=0",
                //    "WpaPskAuth5g=0",
                //    "Wpa2Auth5g=0",
                //    "Wpa2PskAuth5g=0",
                //    "WpaEncryption5g=0",
                //    "WpaPreSharedKey5g=305100001",
                //    "ShowWpaKey5g=0x01",
                //    "RadiusServer5g=0.0.0.0",
                //    "RadiusPort5g=1812",
                //    "RadiusKey5g=",
                //    "WpaRekeyInterval5g=0",
                //    "WpaReauthInterval5g=3600",
                //    "SharedKeyAuthentication5g=0",
                //    "w802_1xAuthentication5g=0",
                //    "Network5gKey1=",
                //    "Network5gKey2=",
                //    "Network5gKey3=",
                //    "Network5gKey4=",
                //    "DefaultSecretKey5g=1",
                //    "WepPassPhrase5g=",
                //    "GenerateWepKeys=0",
                //    "WepKeysGenerated5g=0",
                //    "commitwlan5gPrimaryNetwork=0",
                //    "AutoSecurity5g=1"
                //};

                string[] _wlan5gPrimaryNetworkData = new string[]
                {
                    "PrimaryNetworkEnable5g=1",
                    _ssid5g,
                    "ClosedNetwork5g=0",
                    "BssModeRequired5g=0",
                    "ApIsolate5g=0",
                    "WpaAuth5g=0",
                    "WpaPskAuth5g=0",
                    "Wpa2Auth5g=0",
                    _wpa2pskauth5g,
                    _wpa2enc5g,
                    _wpa2key5g,
                    "ShowWpaKey5g=0x01",
                    "RadiusServer5g=0.0.0.0",
                    "RadiusPort5g=1812",
                    "RadiusKey5g=",
                    "WpaRekeyInterval5g=0",
                    "WpaReauthInterval5g=3600",
                    "SharedKeyAuthentication5g=0",
                    "w802_1xAuthentication5g=0",
                    "Network5gKey1=",
                    "Network5gKey2=",
                    "Network5gKey3=",
                    "Network5gKey4=",
                    "DefaultSecretKey5g=1",
                    "WepPassPhrase5g=",
                    "GenerateWepKeys=0",
                    "WepKeysGenerated5g=0",
                    "commitwlan5gPrimaryNetwork=0",
                    "AutoSecurity5g=1"
                };
 
                _data = string.Empty;

                for (int i = 0; i < _wlan5gPrimaryNetworkData.Length - 1; i++)
                {
                    _data += _wlan5gPrimaryNetworkData[i] + "&";
                }

                _data += _wlan5gPrimaryNetworkData[_wlan5gPrimaryNetworkData.Length - 1];
                //_data = _outputPower + _band;
                //_data = "OutputPower=100&Band=2&NMode=1&NBandwidth=20&ChannelNumber=10&ObssCoexistence=1&STBCTx=0&restoreWirelessDefaults=0&commitwlanRadio=0&scanActions=0";
                //_agent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36";
                _contentType = "application/x-www-form-urlencoded";

                _keepAlive = true;
                WebHttpPostData(_uri, _data, _agent, _contentType, _keepAlive, _Type, _Auth, _username, _password, ref response);
                Thread.Sleep(3000);






            }



            return true;
        }

        public bool TCHLogout(string hostIP)
        {
            /* Logout proccess */
            string response = string.Empty;

            /* Login in process */
            string _uri = @"http://" + hostIP;
            string _agent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; InfoPath.3)";
            string _contentType = "application/x-www-form-urlencoded";
            bool _keepAlive = true;
            bool _Auth = false;
            string _username = "";
            string _password = "";
            string _Type = "";
            string _CSRFValue = string.Empty;       


            /* Get CSRFVaule */
            
            WebHttpGetData(_uri, _agent, _keepAlive, _Auth, _username, _password, "", ref response);


            int index = response.ToLower().IndexOf("csrfvalue");
            if (index == -1)
            {
                Debug.WriteLine("CSRFValue not found");
                return false;
            }

            string _csrfValuetmp = response.Substring(index + 9, 30);
            index = _csrfValuetmp.ToLower().IndexOf("value=");

            if (index == -1)
            {
                Debug.WriteLine("Value not found");
                return false;
            }

            _csrfValuetmp = _csrfValuetmp.Substring(index + 6, 10);

            string[] str = _csrfValuetmp.Split('>');

            _CSRFValue = str[0];

            _uri = @"http://" + hostIP + @"/goform/login";
            string _data = "CSRFValue=" + _CSRFValue + "&logoffUser=1";
            _agent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; InfoPath.3)";
            _contentType = "application/x-www-form-urlencoded";

            _keepAlive = true;
            _Auth = false;
            _username = "";
            _password = "";
            _Type = "";
            try
            {
                WebHttpPostData(_uri, _data, _agent, _contentType, _keepAlive, _Type, _Auth, _username, _password, ref response);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Login Fail: " + ex.ToString());
                return false;
            }
            return true;
              
        }

    }
}



