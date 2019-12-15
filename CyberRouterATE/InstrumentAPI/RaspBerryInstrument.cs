///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang. 
///  File           : RaspBerryInstrument.cs
///  Update         : 2016-09-08    
///  Version        : 1.0.0.20160908
///  Description    : Changes to this file may cause incorrect behavior and will be lost
///                   if the code is regenerated.
///  Modified       : 2016-09-08 Initial version.///                  
///---------------------------------------------------------------------------------------

using System;
using System.Text;
using System.Diagnostics;
using ComportClass;
using System.Threading;
using ComportClass;
using System.Threading;

namespace RaspBerryInstruments
{
    /// <summary>
    /// Turn Table
    /// </summary>
    class turntable
    {
        /* Declare public member variable */

        /* Declare private member variable */        
        private const bool ATE_FAIL = false;
        private const bool ATE_PASS = true;

        private Comport comport;
        
        //馬達正轉指令 需要再加入角度
        //public static string cmd_Motor_CW = "sudo python TTMotor.py b ";
        //static string cmd_Motor_CW = "sudo python TTMotor.py b ";
        static string cmdClockWiseHead = "sudo python TTMotor.py b ";
        //馬達反轉指令 需要再加入角度
        //public static string cmd_Motor_CCW = "sudo python TTMotor.py a ";
        //static string cmd_Motor_CCW = "sudo python TTMotor.py a ";
        static string cmdCounterClockWiseHead = "sudo python TTMotor.py a ";

        public turntable(string portName)
        {
            Comport port = new Comport();
            port.init(portName, TurntableConstants.TurntableBaudrate, TurntableConstants.TurntableParity, TurntableConstants.TurntabletDataBits, TurntableConstants.TurntableStopBits, "none", 500, 500);
            comport = port;
        }

        public turntable(Comport ComX)
        {
            if(ComX != null)
                comport = ComX;
        }

        public bool Login(string username, string password, ref string strRead)
        {
            if (!comport.isOpen()) 
                return false;
            else
            {
                string str = string.Empty;
                strRead = string.Empty;

                comport.DiscardBuffer();
                comport.Write("\r\n");
                Thread.Sleep(200);
                int count = comport.GetBytesToRead();
                str = comport.Read(0, count);
                strRead += str;
                if (str.IndexOf("@") >= 0) return true;
                if (str.IndexOf("login:") < 0) return false;
                
                comport.WriteLine(username);
                Thread.Sleep(200);

                count = comport.GetBytesToRead();
                str = comport.Read(0, count);
                strRead += str;

                if(str.IndexOf("Password:") < 0 ) return false ;
                
                comport.WriteLine(password);
                Thread.Sleep(500);
                count = comport.GetBytesToRead();
                str = comport.Read(0, count);
                strRead += str;

                if(str.IndexOf("Last login") < 0) return false;

                Thread.Sleep(200);
                count = comport.GetBytesToRead();
                str = comport.Read(0, count);
                strRead += str;
            }
            return true;
        }

        public bool Logout(ref string strRead)
        {
            if (!comport.isOpen())
                return false;
            else
            {
                string str = string.Empty;              

                comport.WriteLine("logout");
                Thread.Sleep(200);
                int count = comport.GetBytesToRead();
                str = comport.Read(0, count);           
                
                strRead = str;
            }
            return true;
        }

        public bool ClockWiseTurn(string angle, string calibration, ref string strRead)
        {
            if(!comport.isOpen()) return false;

            string command = cmdClockWiseHead + angle + " " + calibration ;
            comport.WriteLine(command);
            Thread.Sleep(500);
            strRead = comport.ReadLine();
            return true;
        }

        public bool CounterClockWiseTurn(string angle, string calibration, ref string strRead)
        {
            if(!comport.isOpen()) return false;

            string command = cmdCounterClockWiseHead + angle + " " + calibration ;
            comport.WriteLine(command);
            Thread.Sleep(500);
            strRead = comport.ReadLine();
            return true;
        }



        







        /*         
         The microsoft version of enter or new line is \r\n which is 0x0d 0x0a in hex.
            \r is the carriage return
            In a shell or a printer this would put the cursor back to the beginning of the line.
            \n is the line feed
            Puts the cursor one line below, in some shells this also puts the cursor to the beginning of the next line. a printer would simply scroll the paper a bit.
            So much for the history lesson. Current windows systems still use these characters to indicate a line ending. Dos generated this code when pressing enter.
            The key code is a bit different. Beginning with the esc key being the 1. Enter is 28.        
         */








        //public e4438c(string resourceName)
        //{
        //    try
        //    {
        //        mbSessionE4438C = (MessageBasedSession)ResourceManager.GetLocalManager().Open(resourceName);                              
        //    }
        //    catch (InvalidCastException)
        //    {
        //        Debug.WriteLine("Resource selected must be a message-based session.");
        //        mbSessionE4438C = null;
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A00");
        //        mbSessionE4438C = null;
        //    }
        //}

        ///// <summary>
        ///// Use this function call to deallocate the VISA resource.
        ///// </summary>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool Dispose()
        //{
        //    if (mbSessionE4438C != null)
        //        mbSessionE4438C.Dispose();
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// Use this function call to check the resource has been allocated.
        ///// </summary>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool init()
        //{
        //    try
        //    {
        //        if (mbSessionE4438C.Query("*IDN?").IndexOf("E4438C") == -1)
        //            return false;
        //    }
        //    catch (Exception)
        //    {
        //        return false;
        //    }
        //    return true;
        //}

        ///// <summary>
        ///// This function performs PRESET action.
        ///// </summary>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool preset()
        //{
        //    try
        //    {
        //        /* Reset instrument */
        //        mbSessionE4438C.Write("*RST");
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A99");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function performs the following actions:
        ///// - Set Tx frequency with MHz unit. 
        ///// Remote-control command(s):
        ///// [:SOURce]:FREQuency:FIXed
        ///// </summary>
        ///// <param name="Frequency">
        ///// Sets the frequency of RF output.
        ///// Valid Range:
        ///// 250 kHz to 6 GHz
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool configureTxFrequency(double Frequency)
        //{
        //    /* Check frequency range 250 kHz to 6 GHz */
        //    if ((Frequency * 1e6) < agilentConstants.E4438cFrequencyLimitDown || (Frequency * 1e6) > agilentConstants.E4438cFrequencyLimitUp)
        //    {
        //        Debug.WriteLine("Frequency out of range");
        //        Debug.WriteLine("Instrument status error: A01");
        //        return ATE_FAIL;
        //    }

        //    try
        //    {
        //        /* Set center frequency */
        //        Debug.WriteLine("FREQ:FIX " + Frequency + "MHz");
        //        mbSessionE4438C.Write("FREQ:FIX " + Frequency + "MHz");
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A01");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function queries the RF output frequency.
        ///// Remote-control command(s):
        ///// [:SOURce]:FREQuency:FIXed?
        ///// </summary>
        ///// <param name="refFreq">
        ///// This parameter returns the RF output frequency.
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool readTxFrequency(ref string refFreq)
        //{
        //    try
        //    {
        //        /* Query center frequency */
        //        refFreq = mbSessionE4438C.Query("FREQ:FIX?");
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A02");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function configures Tx power level.
        ///// Remote-control command(s):
        ///// [:SOURce]:POWer[:LEVel][:IMMediate][:AMPLitude]
        ///// [:SOURce]:OUTPut:STATe ON|OFF
        ///// :UNIT:POWer DBM
        ///// </summary>
        ///// <param name="Level">
        ///// Set the RF output level.
        ///// Valid Range:
        ///// 25 dBm ~ -136 dBm 
        ///// </param>
        ///// <param name="RF">
        ///// false: RF off
        ///// true : RF on
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool configureTxLevel(double Level, bool RF)
        //{
        //    try
        //    {
        //        /* Level range -136 dBm to 25 dBm */
        //        if (Level < agilentConstants.E4438cTxLevelLimitDown || Level > agilentConstants.E4438cTxLevelLimitUp)
        //        {
        //            Debug.WriteLine("Level out of range");
        //            Debug.WriteLine("Instrument status error: A03");
        //            return ATE_FAIL;
        //        }

        //        mbSessionE4438C.Write("UNIT:POWer DBM");
        //        mbSessionE4438C.Write("POW " + Level);
        //        mbSessionE4438C.Write("OUTPut:STATe " + (RF ? "ON" : "OFF"));
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A03");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function queries the RF output level.
        ///// Remote-control command(s):
        ///// :POWer?
        ///// </summary>
        ///// <param name="refLevel">
        ///// This parameter returns the RF output level (Unit: dBm).
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool readTxLevel(ref string refLevel)
        //{
        //    try
        //    {
        //        /* Query power level */
        //        refLevel = mbSessionE4438C.Query("POW?");
        //        Debug.WriteLine("POWER: " + refLevel);
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A04");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function configures Tx output modulation on/off.
        ///// Remote-control command(s):
        ///// :OUTPut:MODulation[:STATe] ON|OFF|1|0 
        ///// </summary>
        ///// <param name="Modulation">
        ///// Switches the modulator on or off.
        ///// false: OFF
        ///// true: ON
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool configureTxModulation(bool Modulation)
        //{
        //    try
        //    {
        //        /* Modulation output on|off */
        //        mbSessionE4438C.Write("OUTP:MOD " + (Modulation ? "ON" : "OFF"));
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A05");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function configures digital modulation on/off.
        ///// Remote-control command(s):
        ///// [:SOURce]:RADio:DMODulation:ARB[:STATe] ON|OFF|1|0
        ///// </summary>
        ///// <param name="Modulation">
        ///// Switches the digital modulator on or off.
        ///// false: OFF
        ///// true: ON
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool configureDigitalModulation(bool Modulation)
        //{
        //    try
        //    {
        //        /* Digital modulation on|off */
        //        mbSessionE4438C.Write(":RAD:DMOD:ARB " + (Modulation ? "ON" : "OFF"));
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A06");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function only configures multiple carrier on or off. The MCARrier choice selects multicarrier and turns it on. 
        ///// Selecting any other setup such as GSM or CDPD turns multicarrrier off.
        ///// Remote-control command(s):
        ///// [:SOURce]:RADio:DMODulation:ARB:SETup GSM|NADC|PDC|PHS|DECT|AC4Fm|ACQPsk|
        ///// CDPD|PWT|EDGE|TETRa|MCARrier
        ///// </summary>
        ///// <param name="Mode">
        ///// Switches the multicarrier on or off.
        ///// false: OFF
        ///// true: ON
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool configureMulticarrier(bool Mode)
        //{
        //    try
        //    {
        //        /* Multicarrier on */
        //        if (Mode)
        //            mbSessionE4438C.Write("RAD:DMOD:ARB:SET MCAR");
        //        /* Multicarrier off */
        //        else
        //            mbSessionE4438C.Write("RAD:DMOD:ARB:SET NADC");
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A07");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function loads the digital modulation file.
        ///// Remote-control command(s):
        ///// [:SOURce]:RADio:DMODulation:ARB:SETup:MCARrier "file name"
        ///// </summary>
        ///// <param name="FileName">
        ///// Specifies a digital modulation file.
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool loadDModFile(string FileName)
        //{
        //    if (FileName == "")
        //        return ATE_FAIL;

        //    try
        //    {
        //        /* Loads file */
        //        mbSessionE4438C.Write("RAD:DMOD:ARB:SET:MCAR \"" + FileName + "\"");
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A08");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function saves the digital modulation file.
        ///// Remote-control command(s):
        ///// [:SOURce]:RADio:DMODulation:ARB:SETup:MCARrier:STORe "file name"
        ///// </summary>
        ///// <param name="FileName">
        ///// Specifies a digital modulation file.
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool saveDModFile(string FileName)
        //{
        //    if (FileName == "")
        //        return ATE_FAIL;

        //    try
        //    {
        //        /* Saves file */
        //        mbSessionE4438C.Write("RAD:DMOD:ARB:SET:MCAR:STOR \"" + FileName + "\"");
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A09");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function sets the filter alpha value.
        ///// Remote-control command(s):
        ///// [:SOURce]:RADio:DMODulation:ARB:FILTer:ALPHa
        ///// </summary>
        ///// <param name="Value">
        ///// Sets the filter ALPHA value.
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool configureFilterAlpha(double Value)
        //{
        //    /* Alpha range 0.0 to 1.0 */
        //    if (Value < agilentConstants.E4438cFilterAlphaLimitDown || Value > agilentConstants.E4438cFilterAlphaLimitUp)
        //    {
        //        Debug.WriteLine("Alpha out of range");
        //        Debug.WriteLine("Instrument status error: A10");
        //        return ATE_FAIL;
        //    }

        //    try
        //    {
        //        /* Sets the filter alpha value. */
        //        mbSessionE4438C.Write("RAD:DMOD:ARB:FILT:ALPH " + Value);
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A10");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function sets the modulation mode.
        ///// Remote-control command(s):
        ///// [:SOURce]:RADio:DMODulation:ARB:MODulation[:TYPE] BPSK|QPSK|IS95QPSK|
        ///// GRAYQPSK|OQPSK|IS95OQPSK|P4DQPSK|PSK8|PSK16|D8PSK|EDGE|MSK|FSK2|FSK4|FSK8|
        ///// FSK16|C4FM|QAM4|QAM16|QAM32|QAM64|QAM256
        ///// </summary>
        ///// <param name="Mode">
        ///// Sets the modulation mode.
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool configureModulationMode(int Mode)
        //{
        //    /*
        //    string[] modulation = {"BPSK", "QPSK", "IS95QPSK", "GRAYQPSK", "OQPSK", "IS95OQPSK", "P4DQPSK", "PSK8", "PSK16", "D8PSK", "EDGE", "MSK", "FSK2", "FSK4", "FSK8",
        //                            "FSK16", "C4FM", "QAM4", "QAM16", "QAM32", "QAM64", "QAM256"};
        //    */

        //    string[] modulation = {"QAM64", "QAM256", "QPSK"};

        //    try
        //    {
        //        /* Modulation mode, just supported 64QAM, 256QAM and QPSK */
        //        mbSessionE4438C.Write("RAD:DMOD:ARB:MOD " + modulation[Mode]);
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A11");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function sets the symbolrate.
        ///// Remote-control command(s):
        ///// [:SOURce]:RADio:DMODulation:ARB:SRATe
        ///// </summary>
        ///// <param name="Mode">
        ///// Sets the symbolrate with sps unit.
        ///// Valid Range:
        ///// 1Ksps ~ 50Msps
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool configureSymbolRate(double Rate)
        //{
        //    try
        //    {
        //        /* Symbolrate range 1Ksps to 50Msps */
        //        if ((Rate * 1e6) < agilentConstants.E4438cSymbolRateLimitDown || (Rate * 1e6) > agilentConstants.E4438cSymbolRateLimitUp)
        //        {
        //            Debug.WriteLine("Symbolrate out of range");
        //            Debug.WriteLine("Instrument status error: A13");
        //            return ATE_FAIL;
        //        }

        //        /* Sets symbolrate */
        //        mbSessionE4438C.Write("RAD:DMOD:ARB:SRAT " + Rate + "msps");
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A13");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function queries the symbolrate.
        ///// Remote-control command(s):
        ///// [:SOURce]:RADio:DMODulation:ARB:SRATe?
        ///// </summary>
        ///// <param name="refRate">
        ///// This parameter returns the symbolrate (Unit: Ksps).
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool readSymbolRate(ref string refRate)
        //{
        //    try
        //    {
        //        /* Queries symbolrate */
        //        refRate = mbSessionE4438C.Query("RAD:DMOD:ARB:SRAT?");
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A13");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function calibrates IQ DC.
        ///// Remote-control command(s):
        ///// :CALIBration:IQ:DC
        ///// </summary>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool calibrateIQDC()
        //{
        //    try
        //    {
        //        /* Calibrate IQ DC leakage */
        //        mbSessionE4438C.Write(":CALIBration:IQ:DC");
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A14");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}

        ///// <summary>
        ///// This function configures RF output state.
        ///// Remote-control command(s):
        ///// [:SOURce]:OUTPut:STATe ON|OFF
        ///// </summary>
        ///// <param name="State">
        ///// false: RF output off
        ///// true : RF output on
        ///// </param>
        ///// <returns>
        ///// Returns the status code of this operation.
        ///// </returns>
        //public bool configureRFState(bool State)
        //{
        //    try
        //    {
        //        mbSessionE4438C.Write("OUTPut:STATe " + (State ? "ON" : "OFF"));
        //    }
        //    catch (Exception exp)
        //    {
        //        Debug.WriteLine(exp.Message);
        //        Debug.WriteLine("Instrument status error: A15");
        //        return ATE_FAIL;
        //    }
        //    return ATE_PASS;
        //}
    }
    
    public class TurntableConstants
    {
        /* Turntable constant */
        public const int TurntableBaudrate      = 115200;          
        public const int TurntabletDataBits     = 8; 
        public const string TurntableParity     = "None";
        public const string TurntableStopBits   ="one" ;
    }
}
