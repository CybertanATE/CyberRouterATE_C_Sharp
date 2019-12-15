using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CyberRouterATE
{
    public class E8350
    {
        /* Define E8350 CGI command */
        public const string CloseSession = "http://DUTIPaddr/session.cgi?Close_Session=1";
        public const string mfgTest1 = "http://DUTIPaddr/mfgtst.cgi?sys_mfgTest=1";
        public const string B24G_11G = "http://DUTIPaddr/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=11g-only&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=6";
        public const string B24G_11N_M20M = "http://DUTIPaddr/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=6";
        public const string B24G_11N_M40M = "http://DUTIPaddr/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel=6";
        public const string B24G_11N_M20M_Cauto = "http://DUTIPaddr/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=Auto";
        public const string B24G_11N_M40M_Cauto = "http://DUTIPaddr/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel=Auto";
        public const string B5G_11N_M20M = "http://DUTIPaddr/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=an-mixed&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=36";
        public const string B5G_11N_M40M = "http://DUTIPaddr/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=an-mixed&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel=36";
        public const string B5G_11AC_M20M = "";
        public const string B5G_11AC_M40M = "";
        public const string B5G_11AC_M80M = "http://DUTIPaddr/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=ac-mixed&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel=36";
        public const string B5G_11AC_M20M_Cauto = "";
        public const string B5G_11AC_M40M_Cauto = "";
        public const string B5G_11AC_M80M_Cauto = "";
        public const string B5G_11N_M20M_Cauto = "";
        public const string B5G_11N_M40M_Cauto = "http://DUTIPaddr/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=an-mixed&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel36";
        public const string B5G_11N_M80M_Cauto = "";
        
    }

    public class EA8500
    {
        /* Define EA8500 CGI command */
        public const string mfgTest1 = "http://DUTIPaddr:81/mfgtst.cgi?sys_mfgTest=1";
        public const string mfgTest0 = "http://DUTIPaddr:81/mfgtst.cgi?sys_mfgTest=0";
        public const string B24G_11G = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=11g-only&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=6";
        public const string B24G_11N_M20M = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=6";
        public const string B24G_11N_M40M = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel=6";
        public const string B24G_11N_M20M_Cauto = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=Auto";
        public const string B24G_11N_M40M_Cauto = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=11n-2G-only&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel=Auto";
        public const string B5G_11N_M20M = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=an-mixed&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=36";
        public const string B5G_11N_M40M = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=an-mixed&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel=36";
        public const string B5G_11AC_M20M = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=ac-mixed&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=36";
        public const string B5G_11AC_M40M = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=ac-mixed&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel=36";
        public const string B5G_11AC_M80M = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=ac-mixed&sys_BandWidth=80&sys_wlSSID=ssid&sys_wlChannel=36";
        public const string B5G_11AC_M20M_Cauto = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=ac-mixed&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=Auto";
        public const string B5G_11AC_M40M_Cauto = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=ac-mixed&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel=Auto";
        public const string B5G_11AC_M80M_Cauto = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=ac-mixed&sys_BandWidth=80&sys_wlSSID=ssid&sys_wlChannel=Auto";
        public const string B5G_11N_M20M_Cauto = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=an-mixed&sys_BandWidth=20&sys_wlSSID=ssid&sys_wlChannel=Auto";
        public const string B5G_11N_M40M_Cauto = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=an-mixed&sys_BandWidth=40&sys_wlSSID=ssid&sys_wlChannel=Auto";
        public const string B5G_11N_M80M_Cauto = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=an-mixed&sys_BandWidth=80&sys_wlSSID=ssid&sys_wlChannel=Auto";
        public const string B24G_Disabled = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=0&sys_wlMode=disabled";
        public const string B5G_Disabled = "http://DUTIPaddr:81/mfgtst.cgi?sys_wlInterface=1&sys_wlMode=disabled";
    }


    public class EA8500X8
    {
        /* Define EA8500X8 code Command */                
        public const string host = "http://DUTIPaddr/JNAP/";
        //public const string 
        // data_format = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11bg"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":3,""security"":""None""}}]}}]";
        public const string data_2_4G_11bg = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11bg"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_2_4G_11n = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11n"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_2_4G_11g = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11g"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_2_4G_11b = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11b"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_2_4G_11mixed = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_2_4G_Off = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":False,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_2_4G_11mixed_WPAPmixed = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":channel_3,""security"":""WPA-Mixed-Personal"",""wpaPersonalSettings"":{""passphrase"":""password""}}}]}}]";
        public const string data_2_4G_11mixed_WEP = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":channel_3,""security"":""WEP"",""wepSettings"":{""encryption"":""WEP-64"",""key1"":""1234567890"",""key2"":""2345678901"",""key3"":""3456789012"",""key4"":""45678901234"",""txkey"":1}}}]}}]";



        //public const string data_2_4G_11n_20M = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11n"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Standard"",""channel"":channel_3,""security"":""None""}}]}}]";
        //public const string data_2_4G_11n_40M = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11n"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Wide"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_2_4G_11n_20M = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11n"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Standard"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_2_4G_11n_40M = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11n"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":channel_3,""security"":""None""}}]}}]";

        /* channelWidth : Auto , 20M:Standard  40M:Wide */


        public const string dataTemp  = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""jin789"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":3,""security"":""None""}}]}}]";



        public const string data_5G_11n_20M = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11n"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Standard"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_5G_11n_40M = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11n"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Wide"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_5G_11ac_20M = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Standard"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_5G_11ac_40M = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Wide"",""channel"":channel_3,""security"":""None""}}]}}]";
        public const string data_5G_11ac_80M = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":channel_3,""security"":""None""}}]}}]";
        


        public const string data_5G_11mixed = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":3,""security"":""None""}}]}}]";
        public const string data_5G_11an = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11an"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":3,""security"":""None""}}]}}]";
        public const string data_5G_11n = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11n"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":3,""security"":""None""}}]}}]";
        public const string data_5G_11a = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11n"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":3,""security"":""None""}}]}}]";
        public const string data_5G_Off = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":False,""mode"":""802.11n"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":3,""security"":""None""}}]}}]";
        
        public const string dataSrc_2_4G = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":3,""security"":""None""}}]}}]";
        public const string dataSrc_5G_WPA2 = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_5GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":3,""security"":""WPA2-Personal"",""wpaPersonalSettings"":{""passphrase"":""password""}}}]}}]";

        public const string dataSrc_2_4G_WPA2 = @"[{""action"":""http://linksys.com/jnap/wirelessap/SetRadioSettings"",""request"":{""radios"":[{""radioID"":""RADIO_2.4GHz"",""settings"":{""isEnabled"":true,""mode"":""802.11mixed"",""ssid"":""SSIDName"",""broadcastSSID"":true,""channelWidth"":""Auto"",""channel"":3,""security"":""WPA2-Personal"",""wpaPersonalSettings"":{""passphrase"":""password""}}}]}}]";

    }
}
