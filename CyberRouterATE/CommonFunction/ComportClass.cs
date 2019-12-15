using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Diagnostics;

namespace ComportClass
{
    class Comport
    {
        private static SerialPort serialPort;

        public Comport()
        {
            if(serialPort == null)
                serialPort = new SerialPort();
        }

        public bool init(string PortName, int BaudRate, string Parity, int DataBits, string StopBits, string Handshake, int ReadTimeout, int WriteTimeout)
        {
            if (PortName == "") return false;
            SetPortName(PortName);
            SetBaudRate(BaudRate);
            if(!SetParity(Parity)) return false;
            if(!SetDataBits(DataBits)) return false ;
            if(!SetStopBits(StopBits)) return false ;
            if(!SetHandShake(Handshake)) return false ;
            SetReadTimeout(ReadTimeout);
            SetWriteTimeout(WriteTimeout);
            SetPortNewLine("\r\n") ;
            SetDtrEnable(true) ;
            SetRtsEnable(true);
            return true;
        }

        public bool Open()
        {
            if(serialPort.IsOpen == true)
            {
                Debug.WriteLine("Serial port already opened") ;
                return false ;
            }

            try
            {
                serialPort.Open();
            }
            catch (ArgumentOutOfRangeException ArgumentOutOfRangeException)
            {
                Debug.WriteLine("The parameter is incorrect..");
                Debug.WriteLine(ArgumentOutOfRangeException);
                Close();
                return false;
            }
            catch (UnauthorizedAccessException UnauthorizedAccessException)
            {
                Debug.WriteLine("Deny the connection for " + serialPort.PortName);
                Debug.WriteLine(UnauthorizedAccessException);
                Close();
                return false;
            }
            catch (InvalidOperationException InvalidOperationException)
            {
                Debug.WriteLine(serialPort.PortName + " was already using.");
                Debug.WriteLine(InvalidOperationException);
                Close();
                return false;
            }
            catch (System.IO.IOException IOException)
            {
                Debug.WriteLine(serialPort.PortName + " was unavailable!!");
                Debug.WriteLine(IOException);
                Close();
                return false;
            }
            catch (ArgumentException ArgumentException)
            {
                Debug.WriteLine("Illegal string!!");
                Debug.WriteLine(ArgumentException);
                Close();
                return false;
            }
            return true ;

        }
        
        public void Close()
        {
            try
            {
                if(serialPort != null)
                    serialPort.Close();
                Debug.WriteLine("Close");
            }
            catch(Exception ex)
            {
                Debug.WriteLine("Close error: " + ex.ToString());
            }
        }

        public bool Write(string str)
        {
            try
            {
                serialPort.Write(str);
            }
            catch(Exception ex)
            {
                Debug.WriteLine("write error: "+ ex.ToString());
                return false ;
            }
            return true;
        }
         
        public bool WriteLine(string str)
        {
            try
            {
                serialPort.WriteLine(str);
            }
            catch(Exception ex)
            {
                Debug.WriteLine("write line error: " + ex.ToString());
                return false ;
            }
            return true;
        }

        public string ReadLine()
        {
            string data = string.Empty;

            try
            {
                data = serialPort.ReadLine();
                Debug.WriteLine(data);
            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex) ;
            }
            return data ;
        }

        public string Read(int offset, int count)
        {
            string data;
            char[] cRead = new char[count]; 

            try
            {
                serialPort.Read(cRead, offset, count);
                data = new string(cRead);
                data = data.Trim();
                Debug.WriteLine(data);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return null;
            }
            return data;
        }

        public bool DiscardBuffer()
        {
            try
            {
                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
            }
            catch(Exception ex)
            {
                Debug.WriteLine("Discard Buffer error: "+ex.ToString());
                return false ;
            }

            return true ;
        }

        public string SetPortName(string PortName)
        {
            serialPort.PortName = PortName;
            return serialPort.PortName;
        }

        public string GetPortName()
        {
            return serialPort.PortName;
        }

        public int SetBaudRate(int BaudRate)
        {
            serialPort.BaudRate = BaudRate;
            return serialPort.BaudRate;
        }

        public int GetBaudRate()
        {
            return serialPort.BaudRate;
        }

        public bool SetParity(string Parity)
        {
            string s = Parity.ToLower();
            Parity parity = System.IO.Ports.Parity.None;
            switch (s)
            {
                case "none":
                    parity = System.IO.Ports.Parity.None;
                    break;
                case "even":
                    parity = System.IO.Ports.Parity.Even;
                    break;
                case "mark":
                    parity = System.IO.Ports.Parity.Mark;
                    break;
                case "odd":
                    parity = System.IO.Ports.Parity.Odd;
                    break;
                case "space":
                    parity = System.IO.Ports.Parity.Space;
                    break;
                default:
                    return false;
                    break;
            }
            serialPort.Parity = parity;
            return true;
        }

        public Parity GetParity()
        {
            return serialPort.Parity;
        }

        public bool SetDataBits(int DataBits)
        {
            serialPort.DataBits = DataBits;
            return true;
        }

        public int GetDataBits()
        {
            return serialPort.DataBits;
        }

        public bool SetStopBits(string stopBits)
        {
            string s = stopBits.ToLower();
            StopBits stopbits = System.IO.Ports.StopBits.One;
            switch (s)
            { 
                case "one":
                case "1":
                    stopbits = StopBits.One;
                    break;
                case "none":
                    stopbits = StopBits.None;
                    break;
                case "two":
                case "2":
                    stopbits = StopBits.Two;
                    break;
                case "onepointfive":
                case "1.5":
                    stopbits = StopBits.OnePointFive;
                    break;
                default:
                    return false;          
            }
            serialPort.StopBits = stopbits;

            return true;
        }

        public StopBits GetStopBits()
        {
            return serialPort.StopBits;
        }

        public bool SetHandShake(string HandShake)
        {
            string s = HandShake.ToLower();
            Handshake handshake = Handshake.None;

            switch (s)
            {
                case "none":
                    handshake = Handshake.None;
                    break;
                case "xon/xoff":
                    handshake = Handshake.XOnXOff;
                    break;
                case "requestto send":
                    handshake = Handshake.RequestToSend;
                    break;
                case "request to send xonxoff":
                    handshake = Handshake.RequestToSendXOnXOff;
                    break;
                default:
                    return false;
                    break;
            }
            serialPort.Handshake = handshake;
            return true;
        }

        public Handshake GetHandshake()
        {
            return serialPort.Handshake;
        }

        public bool SetReadTimeout(int timeout)
        {
            serialPort.ReadTimeout = timeout;
            return true;
        }

        public int GetReadTimeout()
        {
            return serialPort.ReadTimeout;
        }

        public bool SetWriteTimeout(int timeout)
        {
            serialPort.WriteTimeout = timeout;
            return true;
        }

        public int GetWriteTimeout()
        {
            return serialPort.WriteTimeout;
        }

        public int GetBytesToRead()
        {
            return serialPort.BytesToRead;
        }

        public int GetBytesToWrite()
        {
            return serialPort.BytesToWrite;
        }

        public bool GetCtStatus()
        {
            return serialPort.CDHolding;
        }

        public bool GetDsrStatus()
        {
            return serialPort.DsrHolding;
        }

        public bool SetDtrEnable(bool onoff)
        {
            serialPort.DtrEnable = onoff;
            return true;
        }

        public bool GetDtrEnable()
        {
            return serialPort.DtrEnable;
        }

        public bool SetRtsEnable(bool onoff)
        {
            serialPort.RtsEnable = onoff;
            return true;
        }

        public bool GetRtsEnable()
        {
            return serialPort.RtsEnable;
        }

        public bool SetReadBufferSize(int size)
        {
            serialPort.ReadBufferSize = size;
            return true;
        }

        public int GetReadBufferSize()
        {
            return serialPort.ReadBufferSize;
        }

        public bool SetWriteBufferSize(int size)
        {
            serialPort.WriteBufferSize = size;
            return true;
        }

        public int GetWriteBufferSize()
        {
            return serialPort.WriteBufferSize;
        }

        public bool SetPortNewLine(string newline)
        {
            try
            {
                serialPort.NewLine = newline;
            }
            catch(Exception ex)
            {
                Debug.Write(ex) ;
                return false ;
            }
            return true;
        }

        public string GetNewLine()
        {
            return serialPort.NewLine;
        }

        public bool isOpen()
        {
            return serialPort.IsOpen;
        }









    }

    class Comport2
    {
        private static SerialPort serialPort;

        public Comport2()
        {
            if(serialPort == null)
                serialPort = new SerialPort();
        }

        public bool init(string PortName, int BaudRate, string Parity, int DataBits, string StopBits, string Handshake, int ReadTimeout, int WriteTimeout)
        {
            if (PortName == "") return false;
            SetPortName(PortName);
            SetBaudRate(BaudRate);
            if(!SetParity(Parity)) return false;
            if(!SetDataBits(DataBits)) return false ;
            if(!SetStopBits(StopBits)) return false ;
            if(!SetHandShake(Handshake)) return false ;
            SetReadTimeout(ReadTimeout);
            SetWriteTimeout(WriteTimeout);
            SetPortNewLine("\r\n") ;
            SetDtrEnable(true) ;
            SetRtsEnable(true);
            return true;
        }

        public bool Open()
        {
            if(serialPort.IsOpen == true)
            {
                Debug.WriteLine("Serial port already opened") ;
                return false ;
            }

            try
            {
                serialPort.Open();
            }
            catch (ArgumentOutOfRangeException ArgumentOutOfRangeException)
            {
                Debug.WriteLine("The parameter is incorrect..");
                Debug.WriteLine(ArgumentOutOfRangeException);
                Close();
                return false;
            }
            catch (UnauthorizedAccessException UnauthorizedAccessException)
            {
                Debug.WriteLine("Deny the connection for " + serialPort.PortName);
                Debug.WriteLine(UnauthorizedAccessException);
                Close();
                return false;
            }
            catch (InvalidOperationException InvalidOperationException)
            {
                Debug.WriteLine(serialPort.PortName + " was already using.");
                Debug.WriteLine(InvalidOperationException);
                Close();
                return false;
            }
            catch (System.IO.IOException IOException)
            {
                Debug.WriteLine(serialPort.PortName + " was unavailable!!");
                Debug.WriteLine(IOException);
                Close();
                return false;
            }
            catch (ArgumentException ArgumentException)
            {
                Debug.WriteLine("Illegal string!!");
                Debug.WriteLine(ArgumentException);
                Close();
                return false;
            }
            return true ;

        }
        
        public void Close()
        {
            try
            {
                if(serialPort != null)
                    serialPort.Close();
                Debug.WriteLine("Close");
            }
            catch(Exception ex)
            {
                Debug.WriteLine("Close error: " + ex.ToString());
            }
        }

        public bool Write(string str)
        {
            try
            {
                serialPort.Write(str);
            }
            catch(Exception ex)
            {
                Debug.WriteLine("write error: "+ ex.ToString());
                return false ;
            }
            return true;
        }
         
        public bool WriteLine(string str)
        {
            try
            {
                serialPort.WriteLine(str);
            }
            catch(Exception ex)
            {
                Debug.WriteLine("write line error: " + ex.ToString());
                return false ;
            }
            return true;
        }

        public string ReadLine()
        {
            string data = string.Empty;

            try
            {
                data = serialPort.ReadLine();
                Debug.WriteLine(data);
            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex) ;
            }
            return data ;
        }

        public string Read(int offset, int count)
        {
            string data;
            char[] cRead = new char[count]; 

            try
            {
                serialPort.Read(cRead, offset, count);
                data = new string(cRead);
                data = data.Trim();
                Debug.WriteLine(data);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return null;
            }
            return data;
        }

        public bool DiscardBuffer()
        {
            try
            {
                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
            }
            catch(Exception ex)
            {
                Debug.WriteLine("Discard Buffer error: "+ex.ToString());
                return false ;
            }

            return true ;
        }

        public string SetPortName(string PortName)
        {
            serialPort.PortName = PortName;
            return serialPort.PortName;
        }

        public string GetPortName()
        {
            return serialPort.PortName;
        }

        public int SetBaudRate(int BaudRate)
        {
            serialPort.BaudRate = BaudRate;
            return serialPort.BaudRate;
        }

        public int GetBaudRate()
        {
            return serialPort.BaudRate;
        }

        public bool SetParity(string Parity)
        {
            string s = Parity.ToLower();
            Parity parity = System.IO.Ports.Parity.None;
            switch (s)
            {
                case "none":
                    parity = System.IO.Ports.Parity.None;
                    break;
                case "even":
                    parity = System.IO.Ports.Parity.Even;
                    break;
                case "mark":
                    parity = System.IO.Ports.Parity.Mark;
                    break;
                case "odd":
                    parity = System.IO.Ports.Parity.Odd;
                    break;
                case "space":
                    parity = System.IO.Ports.Parity.Space;
                    break;
                default:
                    return false;
                    break;
            }
            serialPort.Parity = parity;
            return true;
        }

        public Parity GetParity()
        {
            return serialPort.Parity;
        }

        public bool SetDataBits(int DataBits)
        {
            serialPort.DataBits = DataBits;
            return true;
        }

        public int GetDataBits()
        {
            return serialPort.DataBits;
        }

        public bool SetStopBits(string stopBits)
        {
            string s = stopBits.ToLower();
            StopBits stopbits = System.IO.Ports.StopBits.One;
            switch (s)
            { 
                case "one":
                case "1":
                    stopbits = StopBits.One;
                    break;
                case "none":
                    stopbits = StopBits.None;
                    break;
                case "two":
                case "2":
                    stopbits = StopBits.Two;
                    break;
                case "onepointfive":
                case "1.5":
                    stopbits = StopBits.OnePointFive;
                    break;
                default:
                    return false;          
            }
            serialPort.StopBits = stopbits;

            return true;
        }

        public StopBits GetStopBits()
        {
            return serialPort.StopBits;
        }

        public bool SetHandShake(string HandShake)
        {
            string s = HandShake.ToLower();
            Handshake handshake = Handshake.None;

            switch (s)
            {
                case "none":
                    handshake = Handshake.None;
                    break;
                case "xon/xoff":
                    handshake = Handshake.XOnXOff;
                    break;
                case "requestto send":
                    handshake = Handshake.RequestToSend;
                    break;
                case "request to send xonxoff":
                    handshake = Handshake.RequestToSendXOnXOff;
                    break;
                default:
                    return false;
                    break;
            }
            serialPort.Handshake = handshake;
            return true;
        }

        public Handshake GetHandshake()
        {
            return serialPort.Handshake;
        }

        public bool SetReadTimeout(int timeout)
        {
            serialPort.ReadTimeout = timeout;
            return true;
        }

        public int GetReadTimeout()
        {
            return serialPort.ReadTimeout;
        }

        public bool SetWriteTimeout(int timeout)
        {
            serialPort.WriteTimeout = timeout;
            return true;
        }

        public int GetWriteTimeout()
        {
            return serialPort.WriteTimeout;
        }

        public int GetBytesToRead()
        {
            return serialPort.BytesToRead;
        }

        public int GetBytesToWrite()
        {
            return serialPort.BytesToWrite;
        }

        public bool GetCtStatus()
        {
            return serialPort.CDHolding;
        }

        public bool GetDsrStatus()
        {
            return serialPort.DsrHolding;
        }

        public bool SetDtrEnable(bool onoff)
        {
            serialPort.DtrEnable = onoff;
            return true;
        }

        public bool GetDtrEnable()
        {
            return serialPort.DtrEnable;
        }

        public bool SetRtsEnable(bool onoff)
        {
            serialPort.RtsEnable = onoff;
            return true;
        }

        public bool GetRtsEnable()
        {
            return serialPort.RtsEnable;
        }

        public bool SetReadBufferSize(int size)
        {
            serialPort.ReadBufferSize = size;
            return true;
        }

        public int GetReadBufferSize()
        {
            return serialPort.ReadBufferSize;
        }

        public bool SetWriteBufferSize(int size)
        {
            serialPort.WriteBufferSize = size;
            return true;
        }

        public int GetWriteBufferSize()
        {
            return serialPort.WriteBufferSize;
        }

        public bool SetPortNewLine(string newline)
        {
            try
            {
                serialPort.NewLine = newline;
            }
            catch(Exception ex)
            {
                Debug.Write(ex) ;
                return false ;
            }
            return true;
        }

        public string GetNewLine()
        {
            return serialPort.NewLine;
        }

        public bool isOpen()
        {
            return serialPort.IsOpen;
        }
    }


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