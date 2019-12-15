///---------------------------------------------------------------------------------------
///  This code was created by CyberTan James Chu. 
///  File           : agilentinstrument.cs
///  Update         : 2014-08-21    
///  Version        : 1.0.14.0821
///  Description    : Changes to this file may cause incorrect behavior and will be lost
///                   if the code is regenerated.
///  Modified       : 2014-07-15 Initial version.
///                   2014-07-28 Add APIs for Agilent 11713A attenuator/switch driver.
///                   2014-07-29 Modify configureModulationMode() function that 
///                              to support QPSK modulation.
///                   2014-07-30 Modify configureMulticarrierOn() function that to support
///                              multicarrier off. Change configureMulticarrierOn() function 
///                              name to configureMulticarrier().
///                   2014-08-21 Modify Agilent11713A instrument configureAttenuator() function.(Blue)
///---------------------------------------------------------------------------------------

using System;
using System.Text;
using System.Diagnostics;
using Ivi.Visa.Interop;
//using NationalInstruments.VisaNS;

namespace AgilentInstruments
{
    /// <summary>
    /// Agilent E4438C vector signal generator instrument
    /// </summary>
    class e4438c
    {
        /* Declare public member variable */

        /* Declare private member variable */
        private Ivi.Visa.Interop.ResourceManager mbSessionE4438C;
        private FormattedIO488 src;
        //private MessageBasedSession mbSessionE4438C;
        private const bool ATE_FAIL = false;
        private const bool ATE_PASS = true;

        /// <summary>
        /// This function performs the following initialization actions:
        /// - Creates a new instrument driver session.
        /// - Opens a session to the specified device using the interface and address you specify for the Resource Name parameter.
        /// Note:
        /// This function creates a new session each time you invoke it.
        /// </summary>
        /// <param name="resourceName"></param>
        /// This control specifies the interface and address of the device that is to be initialized (Instrument Descriptor). 
        /// The exact grammar to be used in this control is shown in the note below. 
        /// Default Value:  "TCPIP::192.168.67.60::INSTR"
        /// Notes:
        /// (1) Based on the Instrument Descriptor, this operation establishes a communication session with a device.  
        /// The grammar for the Instrument Descriptor is shown below.
        /// Interface   Grammar
        /// ------------------------------------------------------
        /// VXI-11      TCPIP::remote_host::INSTR
        /// GPIB        GPIB[board]::primary address[::secondary address][::INSTR]
        /// The TCPIP keyword is used for VXI-11 interface.
        /// Examples:
        /// (1) VXI-11 - "TCPIP::192.168.67.60::INSTR"
        /// (2) GPIB   - "GPIB::8::INSTR"
        /// 
        public e4438c(string resourceName)
        {
            try
            {
                mbSessionE4438C = new Ivi.Visa.Interop.ResourceManager();
                src = new FormattedIO488();
                src.IO = (IMessage)mbSessionE4438C.Open(resourceName);                          
            }
            catch (InvalidCastException)
            {
                Debug.WriteLine("Resource selected must be a message-based session.");
                mbSessionE4438C = null;
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A00");
                mbSessionE4438C = null;
            }
        }

        /// <summary>
        /// Use this function call to deallocate the VISA resource.
        /// </summary>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool Dispose()
        {
            //if (mbSessionE4438C != null)
            //    mbSessionE4438C..Dispose();
            return ATE_PASS;
        }

        /// <summary>
        /// Use this function call to check the resource has been allocated.
        /// </summary>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool init()
        {
            try
            {
                src.WriteString("*IDN?", true);
                string str = src.ReadString();
                if (str.IndexOf("E4438C") == -1)
                    return false;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// This function performs PRESET action.
        /// </summary>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool preset()
        {
            try
            {
                /* Reset instrument */
                src.WriteString("*RST", true);
                //mbSessionE4438C.Write("*RST");
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A99");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function performs the following actions:
        /// - Set Tx frequency with MHz unit. 
        /// Remote-control command(s):
        /// [:SOURce]:FREQuency:FIXed
        /// </summary>
        /// <param name="Frequency">
        /// Sets the frequency of RF output.
        /// Valid Range:
        /// 250 kHz to 6 GHz
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool configureTxFrequency(double Frequency)
        {
            /* Check frequency range 250 kHz to 6 GHz */
            if ((Frequency * 1e6) < agilentConstants.E4438cFrequencyLimitDown || (Frequency * 1e6) > agilentConstants.E4438cFrequencyLimitUp)
            {
                Debug.WriteLine("Frequency out of range");
                Debug.WriteLine("Instrument status error: A01");
                return ATE_FAIL;
            }

            try
            {
                /* Set center frequency */
                Debug.WriteLine("FREQ:FIX " + Frequency + "MHz");
                src.WriteString("FREQ:FIX " + Frequency + "MHz", true);
                //mbSessionE4438C.Write("FREQ:FIX " + Frequency + "MHz");
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A01");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function queries the RF output frequency.
        /// Remote-control command(s):
        /// [:SOURce]:FREQuency:FIXed?
        /// </summary>
        /// <param name="refFreq">
        /// This parameter returns the RF output frequency.
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool readTxFrequency(ref string refFreq)
        {
            try
            {
                /* Query center frequency */
                src.WriteString("FREQ:FIX?", true);
                //refFreq = mbSessionE4438C.Query("FREQ:FIX?");
                refFreq = src.ReadString();
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A02");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function configures Tx power level.
        /// Remote-control command(s):
        /// [:SOURce]:POWer[:LEVel][:IMMediate][:AMPLitude]
        /// [:SOURce]:OUTPut:STATe ON|OFF
        /// :UNIT:POWer DBM
        /// </summary>
        /// <param name="Level">
        /// Set the RF output level.
        /// Valid Range:
        /// 25 dBm ~ -136 dBm 
        /// </param>
        /// <param name="RF">
        /// false: RF off
        /// true : RF on
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool configureTxLevel(double Level, bool RF)
        {
            try
            {
                /* Level range -136 dBm to 25 dBm */
                if (Level < agilentConstants.E4438cTxLevelLimitDown || Level > agilentConstants.E4438cTxLevelLimitUp)
                {
                    Debug.WriteLine("Level out of range");
                    Debug.WriteLine("Instrument status error: A03");
                    return ATE_FAIL;
                }

                src.WriteString("UNIT:POWer DBM", true);
                src.WriteString("POW " + Level, true);
                src.WriteString("OUTPut:STATe " + (RF ? "ON" : "OFF"), true);
                //mbSessionE4438C.Write("UNIT:POWer DBM");
                //mbSessionE4438C.Write("POW " + Level);
                //mbSessionE4438C.Write("OUTPut:STATe " + (RF ? "ON" : "OFF"));
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A03");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function queries the RF output level.
        /// Remote-control command(s):
        /// :POWer?
        /// </summary>
        /// <param name="refLevel">
        /// This parameter returns the RF output level (Unit: dBm).
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool readTxLevel(ref string refLevel)
        {
            try
            {
                /* Query power level */
                src.WriteString("POW?", true);
                //refLevel = mbSessionE4438C.Query("POW?");
                refLevel = src.ReadString();
                Debug.WriteLine("POWER: " + refLevel);
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A04");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function configures Tx output modulation on/off.
        /// Remote-control command(s):
        /// :OUTPut:MODulation[:STATe] ON|OFF|1|0 
        /// </summary>
        /// <param name="Modulation">
        /// Switches the modulator on or off.
        /// false: OFF
        /// true: ON
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool configureTxModulation(bool Modulation)
        {
            try
            {
                /* Modulation output on|off */
                src.WriteString("OUTP:MOD " + (Modulation ? "ON" : "OFF"), true);
                //mbSessionE4438C.Write("OUTP:MOD " + (Modulation ? "ON" : "OFF"));
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A05");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function configures digital modulation on/off.
        /// Remote-control command(s):
        /// [:SOURce]:RADio:DMODulation:ARB[:STATe] ON|OFF|1|0
        /// </summary>
        /// <param name="Modulation">
        /// Switches the digital modulator on or off.
        /// false: OFF
        /// true: ON
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool configureDigitalModulation(bool Modulation)
        {
            try
            {
                /* Digital modulation on|off */
                src.WriteString(":RAD:DMOD:ARB " + (Modulation ? "ON" : "OFF"), true);
                //mbSessionE4438C.Write(":RAD:DMOD:ARB " + (Modulation ? "ON" : "OFF"));
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A06");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function only configures multiple carrier on or off. The MCARrier choice selects multicarrier and turns it on. 
        /// Selecting any other setup such as GSM or CDPD turns multicarrrier off.
        /// Remote-control command(s):
        /// [:SOURce]:RADio:DMODulation:ARB:SETup GSM|NADC|PDC|PHS|DECT|AC4Fm|ACQPsk|
        /// CDPD|PWT|EDGE|TETRa|MCARrier
        /// </summary>
        /// <param name="Mode">
        /// Switches the multicarrier on or off.
        /// false: OFF
        /// true: ON
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool configureMulticarrier(bool Mode)
        {
            try
            {
                /* Multicarrier on */
                if (Mode)
                    src.WriteString("RAD:DMOD:ARB:SET MCAR", true);
                    //mbSessionE4438C.Write("RAD:DMOD:ARB:SET MCAR");
                /* Multicarrier off */
                else
                    src.WriteString("RAD:DMOD:ARB:SET NADC", true);
                    //mbSessionE4438C.Write("RAD:DMOD:ARB:SET NADC");
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A07");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function loads the digital modulation file.
        /// Remote-control command(s):
        /// [:SOURce]:RADio:DMODulation:ARB:SETup:MCARrier "file name"
        /// </summary>
        /// <param name="FileName">
        /// Specifies a digital modulation file.
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool loadDModFile(string FileName)
        {
            if (FileName == "")
                return ATE_FAIL;

            try
            {
                /* Loads file */
                src.WriteString("RAD:DMOD:ARB:SET:MCAR \"" + FileName + "\"", true);
                //mbSessionE4438C.Write("RAD:DMOD:ARB:SET:MCAR \"" + FileName + "\"");
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A08");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function saves the digital modulation file.
        /// Remote-control command(s):
        /// [:SOURce]:RADio:DMODulation:ARB:SETup:MCARrier:STORe "file name"
        /// </summary>
        /// <param name="FileName">
        /// Specifies a digital modulation file.
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool saveDModFile(string FileName)
        {
            if (FileName == "")
                return ATE_FAIL;

            try
            {
                /* Saves file */
                src.WriteString("RAD:DMOD:ARB:SET:MCAR:STOR \"" + FileName + "\"", true);
                //mbSessionE4438C.Write("RAD:DMOD:ARB:SET:MCAR:STOR \"" + FileName + "\"");
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A09");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function sets the filter alpha value.
        /// Remote-control command(s):
        /// [:SOURce]:RADio:DMODulation:ARB:FILTer:ALPHa
        /// </summary>
        /// <param name="Value">
        /// Sets the filter ALPHA value.
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool configureFilterAlpha(double Value)
        {
            /* Alpha range 0.0 to 1.0 */
            if (Value < agilentConstants.E4438cFilterAlphaLimitDown || Value > agilentConstants.E4438cFilterAlphaLimitUp)
            {
                Debug.WriteLine("Alpha out of range");
                Debug.WriteLine("Instrument status error: A10");
                return ATE_FAIL;
            }

            try
            {
                /* Sets the filter alpha value. */
                src.WriteString("RAD:DMOD:ARB:FILT:ALPH " + Value, true);
                //mbSessionE4438C.Write("RAD:DMOD:ARB:FILT:ALPH " + Value);
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A10");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function sets the modulation mode.
        /// Remote-control command(s):
        /// [:SOURce]:RADio:DMODulation:ARB:MODulation[:TYPE] BPSK|QPSK|IS95QPSK|
        /// GRAYQPSK|OQPSK|IS95OQPSK|P4DQPSK|PSK8|PSK16|D8PSK|EDGE|MSK|FSK2|FSK4|FSK8|
        /// FSK16|C4FM|QAM4|QAM16|QAM32|QAM64|QAM256
        /// </summary>
        /// <param name="Mode">
        /// Sets the modulation mode.
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool configureModulationMode(int Mode)
        {
            /*
            string[] modulation = {"BPSK", "QPSK", "IS95QPSK", "GRAYQPSK", "OQPSK", "IS95OQPSK", "P4DQPSK", "PSK8", "PSK16", "D8PSK", "EDGE", "MSK", "FSK2", "FSK4", "FSK8",
                                    "FSK16", "C4FM", "QAM4", "QAM16", "QAM32", "QAM64", "QAM256"};
            */

            string[] modulation = {"QAM64", "QAM256", "QPSK"};

            try
            {
                /* Modulation mode, just supported 64QAM, 256QAM and QPSK */
                src.WriteString("RAD:DMOD:ARB:MOD " + modulation[Mode], true);
                //mbSessionE4438C.Write("RAD:DMOD:ARB:MOD " + modulation[Mode]);
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A11");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function sets the symbolrate.
        /// Remote-control command(s):
        /// [:SOURce]:RADio:DMODulation:ARB:SRATe
        /// </summary>
        /// <param name="Mode">
        /// Sets the symbolrate with sps unit.
        /// Valid Range:
        /// 1Ksps ~ 50Msps
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool configureSymbolRate(double Rate)
        {
            try
            {
                /* Symbolrate range 1Ksps to 50Msps */
                if ((Rate * 1e6) < agilentConstants.E4438cSymbolRateLimitDown || (Rate * 1e6) > agilentConstants.E4438cSymbolRateLimitUp)
                {
                    Debug.WriteLine("Symbolrate out of range");
                    Debug.WriteLine("Instrument status error: A13");
                    return ATE_FAIL;
                }

                /* Sets symbolrate */
                src.WriteString("RAD:DMOD:ARB:SRAT " + Rate + "msps", true);
                //mbSessionE4438C.Write("RAD:DMOD:ARB:SRAT " + Rate + "msps");
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A13");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function queries the symbolrate.
        /// Remote-control command(s):
        /// [:SOURce]:RADio:DMODulation:ARB:SRATe?
        /// </summary>
        /// <param name="refRate">
        /// This parameter returns the symbolrate (Unit: Ksps).
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool readSymbolRate(ref string refRate)
        {
            try
            {
                /* Queries symbolrate */
                src.WriteString("RAD:DMOD:ARB:SRAT?", true);
                //refRate = mbSessionE4438C.Query("RAD:DMOD:ARB:SRAT?");
                refRate = src.ReadString();
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A13");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function calibrates IQ DC.
        /// Remote-control command(s):
        /// :CALIBration:IQ:DC
        /// </summary>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool calibrateIQDC()
        {
            try
            {
                /* Calibrate IQ DC leakage */
                src.WriteString(":CALIBration:IQ:DC", true);
                //mbSessionE4438C.Write(":CALIBration:IQ:DC");
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A14");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function configures RF output state.
        /// Remote-control command(s):
        /// [:SOURce]:OUTPut:STATe ON|OFF
        /// </summary>
        /// <param name="State">
        /// false: RF output off
        /// true : RF output on
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool configureRFState(bool State)
        {
            try
            {
                src.WriteString("OUTPut:STATe " + (State ? "ON" : "OFF"), true);
                //mbSessionE4438C.Write("OUTPut:STATe " + (State ? "ON" : "OFF"));
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: A15");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }
    }

    /// <summary>
    /// Agilent 11713A attenuator/switch instrument
    /// </summary>
    class Agilent11713A
    {
        /* Declare public member variable */

        /* Declare private member variable */
        private Ivi.Visa.Interop.ResourceManager mbSession11713A;
        private FormattedIO488 src;
        //private MessageBasedSession mbSession11713A;
        private const bool ATE_FAIL = false;
        private const bool ATE_PASS = true;

        /// <summary>
        /// This function performs the following initialization actions:
        /// - Creates a new instrument driver session.
        /// - Opens a session to the specified device using the interface and address you specify for the Resource Name parameter.
        /// Note:
        /// This function creates a new session each time you invoke it.
        /// </summary>
        /// <param name="resourceName"></param>
        /// This control specifies the interface and address of the device that is to be initialized (Instrument Descriptor). 
        /// The exact grammar to be used in this control is shown in the note below. 
        /// Default Value:  "TCPIP::192.168.67.60::INSTR"
        /// Notes:
        /// (1) Based on the Instrument Descriptor, this operation establishes a communication session with a device.  
        /// The grammar for the Instrument Descriptor is shown below.
        /// Interface   Grammar
        /// ------------------------------------------------------
        /// VXI-11      TCPIP::remote_host::INSTR
        /// GPIB        GPIB[board]::primary address[::secondary address][::INSTR]
        /// The TCPIP keyword is used for VXI-11 interface.
        /// Examples:
        /// (1) VXI-11 - "TCPIP::192.168.67.60::INSTR"
        /// (2) GPIB   - "GPIB::8::INSTR"
        /// 
        public Agilent11713A(string resourceName)
        {
            try
            {
                //mbSession11713A = (MessageBasedSession)ResourceManager.GetLocalManager().Open(resourceName);
                //mbSession11713A = ResourceManager.GetLocalManager().Open(resourceName);
                mbSession11713A = new Ivi.Visa.Interop.ResourceManager();
                src = new FormattedIO488();
                src.IO = (IMessage)mbSession11713A.Open(resourceName);
                
            }
            catch (InvalidCastException)
            {
                Debug.WriteLine("Resource selected must be a message-based session.");
                mbSession11713A = null;
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: G00");
                mbSession11713A = null;
            }
        }

        /// <summary>
        /// Use this function call to deallocate the VISA resource.
        /// </summary>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool Dispose()
        {
            //if (mbSession11713A != null)
            //    mbSession11713A.Dispose();
            return ATE_PASS;
        }

        /// <summary>
        /// This function performs PRESET action.
        /// </summary>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool preset()
        {
            try
            {
                /* Reset instrument */
                src.WriteString("B12345678",true);
                //mbSession11713A.Write("B12345678");
            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: G99");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        /// <summary>
        /// This function configures attenuator action.
        /// </summary>
        /// <param name="Value">
        /// Sets the attenuator value.
        /// </param>
        /// <returns>
        /// Returns the status code of this operation.
        /// </returns>
        public bool configureAttenuator(int Value)
        {
            try
            {
                /* Attenuator range 0dB to 121dB */
                if (Value < agilentConstants.Agilent11713AAttenuatorLimitDown || Value > agilentConstants.Agilent11713AAttenuatorLimitUp)
                {
                    Debug.WriteLine("Attenuator out of range");
                    Debug.WriteLine("Instrument status error: G01");
                    return ATE_FAIL;
                }

                /* Reset instrument */
                src.WriteString("B12345678", true);
                //mbSession11713A.Write("B12345678");
                int Value10 = Value;

                if (Value10 >= 40)   //40
                {
                    src.WriteString("A7", true);
                    //mbSession11713A.Write("A7");
                    Value10 = Value10 - 40;
                }

                if (Value10 >= 40)   //40
                {
                    src.WriteString("A8", true);
                    //mbSession11713A.Write("A8");
                    Value10 = Value10 - 40;
                }

                if (Value10 >= 20)   //20
                {
                    src.WriteString("A6", true);
                    //mbSession11713A.Write("A6");
                    Value10 = Value10 - 20;
                }

                if (Value10 >= 10)   //10
                {
                    src.WriteString("A5", true);
                    //mbSession11713A.Write("A5");
                    Value10 = Value10 - 10;
                }

               // int Value1 = Value % 10;

                if (Value10 >= 4)    //4
                {
                    src.WriteString("A3", true);
                    //mbSession11713A.Write("A3");
                    Value10 = Value10 - 4;
                }

                if (Value10 >= 4)    //4
                {
                    src.WriteString("A4", true);
                    //mbSession11713A.Write("A4");
                    Value10 = Value10 - 4;
                }

                if (Value10 >= 2)    //2
                {
                    src.WriteString("A2", true);
                    //mbSession11713A.Write("A2");
                    Value10 = Value10 - 2;
                }

                if (Value10 >= 1)    //1
                    src.WriteString("A1", true);
                    //mbSession11713A.Write("A1");

               

            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: G01");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }


        public bool configureAttenuator(string Value)
        {
            try
            {
                /* Reset instrument */
                src.WriteString("B12345678", true);
                //mbSession11713A.Write("B12345678");  
                src.WriteString("A" + Value, true);
                //mbSession11713A.Write("A"+Value);

            }
            catch (Exception exp)
            {
                Debug.WriteLine(exp.Message);
                Debug.WriteLine("Instrument status error: G01");
                return ATE_FAIL;
            }
            return ATE_PASS;
        }


    }

    public class agilentConstants
    {
        /* E4438C constant */
        public const double E4438cFrequencyLimitUp          = (double)6E9;
        public const double E4438cFrequencyLimitDown        = (double)250E3;
        public const int E4438cTxLevelLimitUp               = 25;           // 25dBm
        public const int E4438cTxLevelLimitDown             = -136;         // -136dBm
        public const double E4438cFilterAlphaLimitUp        = 1;            // 1
        public const double E4438cFilterAlphaLimitDown      = 0;            // 0
        public const double E4438cSymbolRateLimitUp         = (double)50E6; // 50Msps
        public const double E4438cSymbolRateLimitDown       = 1000;         // 1Ksps

        /* 11713A constant */
        public const int Agilent11713AAttenuatorLimitUp     = 121;          // 121dB
        public const int Agilent11713AAttenuatorLimitDown   = 0;            // 0dB
    }
}
