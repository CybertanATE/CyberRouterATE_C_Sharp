using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Diagnostics;


namespace Ns_CbtFtpClient
{
    class CbtFtpClient
    {
        private FtpWebRequest _ftpRequest = null;
        private FtpWebResponse _ftpResponse = null;
        private string _ftpPrefix = "ftp://";
        private string _ftpServerIP = string.Empty;
        private string _ftpUserID = string.Empty;
        private string _ftpUserPWD = string.Empty;
        private string _ftpFileName = string.Empty;
        private string _ftpUrl = string.Empty;
        private int _ftpServerPort;
        private bool _passiveMode = false;

        public bool passiveMode
        {
            get { return this._passiveMode; }
            set { this._passiveMode = value; }
        }

        public CbtFtpClient()
        {
            this._ftpServerPort = 21;
        }

        public CbtFtpClient(string ip, string username, string password)
        {
            Init(ip, username, password);
        }

        public bool Init(string ip, string username, string password)
        {
            _ftpUrl = "ftp://" + ip;
            _ftpUserID = username;
            _ftpUserPWD = password;
            this._ftpServerPort = 21;
            //_ftpRequest = (FtpWebRequest)FtpWebRequest.Create(_ftpUrl);
            return true;
        }

        public bool Login()
        {
            //List<string> sItemList = new List<string>();
            _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(_ftpUrl);
            //if (_ftpUserID.ToLower() == "anonymous")
            //{
            //    //_ftpRequest.Credentials = new NetworkCredential();
            //}
            //else
            //{
            _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
            //}
            _ftpRequest.UsePassive = this._passiveMode;
            
            _ftpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
            try
            {
                _ftpResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                Stream responseStream = _ftpResponse.GetResponseStream();

                _ftpResponse.Close();
                _ftpRequest = null;
                return true;
            }
            /*try
            {
                StreamReader SR = new StreamReader(_ftpRequest.GetResponse().GetResponseStream());
                
                SR.Close();
                return true;
            }*/
            catch (Exception ex)
            {
                if (ex.ToString().IndexOf("550") > 0)
                    return true;
                else
                    return false;
            }
        }
        public bool Login(ref string exceptionInfo)
        {
            //List<string> sItemList = new List<string>();
            _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(_ftpUrl);
            //if (_ftpUserID.ToLower() == "anonymous")
            //{
            //    //_ftpRequest.Credentials = new NetworkCredential();
            //}
            //else
            //{
            _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
            //}
            _ftpRequest.UsePassive = this._passiveMode;

            _ftpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
            try
            {
                _ftpResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                Stream responseStream = _ftpResponse.GetResponseStream();

                _ftpResponse.Close();
                _ftpRequest = null;
                return true;
            }
            /*try
            {
                StreamReader SR = new StreamReader(_ftpRequest.GetResponse().GetResponseStream());
                
                SR.Close();
                return true;
            }*/
            catch (Exception ex)
            {
                exceptionInfo = ex.ToString();
                if (ex.ToString().IndexOf("550") > 0)
                    return true;
                else
                    return false;
            }
        }

        public int ftpServerPort
        {
            get
            {
                return _ftpServerPort;
            }

            set
            {
                _ftpServerPort = value;
            }
        }

        public bool UploadFile(string sFilePath, string sFileName, string sSrcFileName, ref string str)
        {
            try
            {
                string sTotalFilePath = _ftpUrl + sFilePath + sFileName;
                //_ftpFilePath = _ftpUrl + sFilePath + sFileName;

                // Get the object used to communicate with the server.
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalFilePath);
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;

                //_ftpRequest.Method = WebRequestMethods.File.UploadFile;

                // This example assumes the FTP site uses anonymous logon. 
                NetworkCredential ftpCredential = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.Credentials = ftpCredential;

                // Copy the contents of the file to the request stream.  
                StreamReader sourceStream = new StreamReader(sSrcFileName);
                byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                sourceStream.Close();
                _ftpRequest.ContentLength = fileContents.Length;

                Stream requestStream = _ftpRequest.GetRequestStream();
                requestStream.Write(fileContents, 0, fileContents.Length);
                requestStream.Close();

                FtpWebResponse response = (FtpWebResponse)_ftpRequest.GetResponse();
                str = response.StatusDescription;
                //Console.WriteLine("Upload File Complete, status {0}", response.StatusDescription);

                response.Close();

                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp upload failed: " + ex.ToString());
                throw new Exception("Ftp upload failed: " + ex.ToString());
                //return false;
            }

            return true;
        }

        public bool DownloadFile(string sFilePath, string sFileName, ref string str)
        {
            try
            {
                //if (sFilePath == "")
                //{
                //    _ftpFilePath = _ftpUrl + "/" + sFilePath + sFileName;
                //}
                //else
                //{
                //    _ftpFilePath = _ftpUrl + sFilePath + sFileName;
                //}

                string sTotalSrcFilePath = _ftpUrl + sFilePath + sFileName;

                //_ftpFilePath = _ftpUrl + sFilePath + sFileName;

                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalSrcFilePath);
                _ftpRequest.UsePassive = this._passiveMode;
                NetworkCredential ftpCredential = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.Credentials = ftpCredential;
                _ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                FtpWebResponse response = (FtpWebResponse)_ftpRequest.GetResponse();


                // Get the object used to communicate with the server.  
                //FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://www.contoso.com/test.htm");
                //request.Method = WebRequestMethods.Ftp.DownloadFile;

                //// This example assumes the FTP site uses anonymous logon.  
                //request.Credentials = new NetworkCredential("anonymous", "janeDoe@contoso.com");

                //FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                Stream responseStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream);

                str = reader.ReadToEnd();
                str += response.StatusDescription;

                //Console.WriteLine(reader.ReadToEnd());

                //Console.WriteLine("Download Complete, status {0}", response.StatusDescription);


                reader.Close();
                response.Close();

                reader.Dispose();

                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp upload failed: " + ex.ToString());
                throw new Exception("Ftp upload failed: " + ex.ToString());
                //return false;
            }

            return true;
        }
        public bool DownloadFile(string sLocalPath, string sFilePath, string sFileName, ref string str)
        {
            FtpWebRequest remoteFileLenReq; // 此請求是為了獲取遠端文件長度
            FtpWebRequest remoteFileReadReq;// 此請求是為了讀取文件
            Stream readStream = null;       // 讀取流
            FileStream writeStream = null;  // 寫本地文件流
            string sLocalFile = sLocalPath + "\\" + sFileName;

            try
            {
                if (File.Exists(sLocalFile))
                {
                    writeStream = new FileStream(sLocalFile, FileMode.Append);

                }
                else
                {
                    writeStream = new FileStream(sLocalPath + "\\" + sFileName, FileMode.Create);
                }

                long startPosition = writeStream.Length;// 讀出本地文件已有長度

                string sTotalSrcFilePath = _ftpPrefix + sFilePath + sFileName;

                // 下面程式碼目的是取遠端文件長度
                //remoteFileLenReq = (FtpWebRequest)FtpWebRequest.Create(sTotalSrcFilePath);
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalSrcFilePath);
                NetworkCredential ftpCredential = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.Credentials = ftpCredential;
                _ftpRequest.UseBinary = true;
                _ftpRequest.ContentOffset = 0;
                _ftpRequest.Method = WebRequestMethods.Ftp.GetFileSize;
                FtpWebResponse rsp = (FtpWebResponse)_ftpRequest.GetResponse();
                long totalByte = rsp.ContentLength;
                rsp.Close();

                if (startPosition >= totalByte)
                {
                    //str += "本地文件長度" + startPosition + "已經大於等於遠端文件長度" + totalByte;
                    str += "File Exists and has been download Finished. Skip the download.";
                    writeStream.Close();

                    return true;
                }

                // 初始化讀取遠端文件請求
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalSrcFilePath);
                //NetworkCredential ftpCredential = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.Credentials = ftpCredential;
                _ftpRequest.UseBinary = true;
                _ftpRequest.KeepAlive = false;
                _ftpRequest.ContentOffset = startPosition;
                _ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                FtpWebResponse response = (FtpWebResponse)_ftpRequest.GetResponse();
                readStream = response.GetResponseStream();

                long downloadedByte = startPosition;
                int bufferSize = 512;
                byte[] btArray = new byte[bufferSize];
                int contentSize = readStream.Read(btArray, 0, btArray.Length);

                while (contentSize > 0)
                {
                    downloadedByte += contentSize;
                    //int percent = (int)(downloadedByte * 100 / totalByte);
                    //System.Console.WriteLine("percent=" + percent + "%");

                    writeStream.Write(btArray, 0, contentSize);
                    contentSize = readStream.Read(btArray, 0, btArray.Length);
                }
                readStream.Close();
                writeStream.Close();
                response.Close();
                return true;





                ////_ftpFilePath = _ftpUrl + sFilePath + sFileName;

                //_ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalFilePath);
                //NetworkCredential ftpCredential = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                //_ftpRequest.Credentials = ftpCredential;
                //_ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                //FtpWebResponse response = (FtpWebResponse)_ftpRequest.GetResponse();


                //// Get the object used to communicate with the server.  
                ////FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://www.contoso.com/test.htm");
                ////request.Method = WebRequestMethods.Ftp.DownloadFile;

                ////// This example assumes the FTP site uses anonymous logon.  
                ////request.Credentials = new NetworkCredential("anonymous", "janeDoe@contoso.com");

                ////FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                //Stream responseStream = response.GetResponseStream();
                ////FileStream writeStream = new FileStream(sLocalPath + "\\" + sFileName, FileMode.Create);

                //int Length = 2048;
                //Byte[] buffer = new Byte[Length];
                //int bytesRead = responseStream.Read(buffer, 0, Length);

                //while (bytesRead > 0)
                //{
                //    writeStream.Write(buffer, 0, bytesRead);
                //    bytesRead = responseStream.Read(buffer, 0, Length);
                //}

                //responseStream.Close();
                //writeStream.Close();

                ////Stream responseStream = response.GetResponseStream();
                ////StreamReader reader = new StreamReader(responseStream);

                ////str = reader.ReadToEnd();
                ////str += response.StatusDescription;

                //////Console.WriteLine(reader.ReadToEnd());

                //////Console.WriteLine("Download Complete, status {0}", response.StatusDescription);


                ////reader.Close();
                //response.Close();

                ////reader.Dispose();

                //_ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp Downloload failed: " + ex.ToString());
                str += "Exception: Ftp Downloload failed: " + ex.ToString();
                throw new Exception("Ftp upload failed: " + ex.ToString());
                //return false;
            }
            finally
            {
                if (readStream != null)
                {
                    readStream.Close();
                }
                if (writeStream != null)
                {
                    writeStream.Close();
                }
            }

            return true;
        }
        

        public bool UploadFileUSBStorage(string sDestination)
        {
            try
            {
                string sTotalDestination = _ftpUrl + sDestination + "/FTPUploadFile.txt";
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalDestination);
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;

                if (_ftpUserID.ToLower() == "anonymous")
                {
                    //_ftpRequest.Credentials = new NetworkCredential();
                }
                else
                {
                    _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                }

                StreamReader sourceStream = new StreamReader("Archives\\USBStorage\\FTP Files\\FTPUploadFile.txt");
                byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                sourceStream.Close();
                _ftpRequest.ContentLength = fileContents.Length;

                Stream requestStream = _ftpRequest.GetRequestStream();
                requestStream.Write(fileContents, 0, fileContents.Length);
                requestStream.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp upload failed: " + ex.ToString());
                return false;
            }

            return true;
        }
        public bool UploadFileUSBStorage(string sDestination, ref string exceptionInfo)
        {
            try
            {
                string sTotalDestination = _ftpUrl + sDestination + "/FTPUploadFile.txt";
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalDestination);
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;

                if (_ftpUserID.ToLower() == "anonymous")
                {
                    //_ftpRequest.Credentials = new NetworkCredential();
                }
                else
                {
                    _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                }

                StreamReader sourceStream = new StreamReader("Archives\\USBStorage\\FTP Files\\FTPUploadFile.txt");
                byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                sourceStream.Close();
                _ftpRequest.ContentLength = fileContents.Length;

                Stream requestStream = _ftpRequest.GetRequestStream();
                requestStream.Write(fileContents, 0, fileContents.Length);
                requestStream.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                exceptionInfo = ex.ToString();
                Debug.WriteLine("Ftp upload failed: " + ex.ToString());
                return false;
            }

            return true;
        }

        public bool DownloadFileUSBStorage(string sSource)
        {
            try
            {
                string sTotalSource = _ftpUrl + sSource + "/FTPDownloadFile.txt";
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalSource);
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                if (_ftpUserID.ToLower() == "anonymous")
                {
                    //_ftpRequest.Credentials = new NetworkCredential();
                }
                else
                {
                    _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                }

                FtpWebResponse sourceResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                FileStream writeStream = null;
                if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Archives\\USBStorage\\FTP Files\\FTPDownloadFile.txt"))
                {
                    writeStream = new FileStream(System.Windows.Forms.Application.StartupPath + "\\Archives\\USBStorage\\FTP Files\\FTPDownloadFile.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    writeStream.Flush();
                    writeStream.SetLength(0);
                }
                else
                {
                    writeStream = new FileStream(System.Windows.Forms.Application.StartupPath + "\\Archives\\USBStorage\\FTP Files\\FTPDownloadFile.txt", FileMode.Create);
                }
                int bufferSize = 512;
                byte[] btArray = new byte[bufferSize];
                int contentSize = sourceResponse.GetResponseStream().Read(btArray, 0, btArray.Length);

                writeStream.Write(btArray, 0, contentSize);
                writeStream.Close();
                sourceResponse.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp Downloload failed: " + ex.ToString());
                return false;
            }

            return true;
        }
        public bool DownloadFileUSBStorage(string sSource, ref string exceptionInfo)
        {
            try
            {
                string sTotalSource = _ftpUrl + sSource + "/FTPDownloadFile.txt";
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalSource);
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                if (_ftpUserID.ToLower() == "anonymous")
                {
                    //_ftpRequest.Credentials = new NetworkCredential();
                }
                else
                {
                    _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                }

                FtpWebResponse sourceResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                FileStream writeStream = null;
                if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Archives\\USBStorage\\FTP Files\\FTPDownloadFile.txt"))
                {
                    writeStream = new FileStream(System.Windows.Forms.Application.StartupPath + "\\Archives\\USBStorage\\FTP Files\\FTPDownloadFile.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    writeStream.Flush();
                    writeStream.SetLength(0);
                }
                else
                {
                    writeStream = new FileStream(System.Windows.Forms.Application.StartupPath + "\\Archives\\USBStorage\\FTP Files\\FTPDownloadFile.txt", FileMode.Create);
                }
                int bufferSize = 512;
                byte[] btArray = new byte[bufferSize];
                int contentSize = sourceResponse.GetResponseStream().Read(btArray, 0, btArray.Length);

                writeStream.Write(btArray, 0, contentSize);
                writeStream.Close();
                sourceResponse.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                exceptionInfo = ex.ToString();
                Debug.WriteLine("Ftp Downloload failed: " + ex.ToString());
                return false;
            }

            return true;
        }

        public bool CreateFolderUSBStorage(string sDestination)
        {
            try
            {
                string sTotalDestination = _ftpUrl + sDestination;
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalDestination);
                if (_ftpUserID.ToLower() == "anonymous")
                {
                    //_ftpRequest.Credentials = new NetworkCredential();
                }
                else
                {
                    _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                }
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.MakeDirectory;

                _ftpResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                _ftpResponse.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp create folder failed: " + ex.ToString());
                return false;
            }

            return true;
        }
        public bool CreateFolderUSBStorage(string sDestination, ref string exceptionInfo)
        {
            try
            {
                string sTotalDestination = _ftpUrl + sDestination;
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalDestination);
                if (_ftpUserID.ToLower() == "anonymous")
                {
                    //_ftpRequest.Credentials = new NetworkCredential();
                }
                else
                {
                    _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                }
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.MakeDirectory;

                _ftpResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                _ftpResponse.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                exceptionInfo = ex.ToString();
                Debug.WriteLine("Ftp create folder failed: " + ex.ToString());
                return false;
            }

            return true;
        }

        public bool DeleteFolderUSBStorage(string sDestination)
        {
            try
            {
                string sTotalDestination = _ftpUrl + sDestination;
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalDestination);
                if (_ftpUserID.ToLower() == "anonymous")
                {
                    //_ftpRequest.Credentials = new NetworkCredential();
                }
                else
                {
                    _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                }
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.RemoveDirectory;

                _ftpResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                _ftpResponse.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp delete folder failed: " + ex.ToString());
                //throw new Exception("Ftp delete folder failed: " + ex.ToString());
                return false;
            }

            return true;
        }
        public bool DeleteFolderUSBStorage(string sDestination, ref string exceptionInfo)
        {
            try
            {
                string sTotalDestination = _ftpUrl + sDestination;
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalDestination);
                if (_ftpUserID.ToLower() == "anonymous")
                {
                    //_ftpRequest.Credentials = new NetworkCredential();
                }
                else
                {
                    _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                }
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.RemoveDirectory;

                _ftpResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                _ftpResponse.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                exceptionInfo = ex.ToString();
                Debug.WriteLine("Ftp delete folder failed: " + ex.ToString());
                //throw new Exception("Ftp delete folder failed: " + ex.ToString());
                return false;
            }

            return true;
        }

        public bool RenameUSBStorage(string sDestination, string sRenameTo)
        {
            try
            {
                string sTotalDestination = _ftpUrl + sDestination;
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalDestination);
                _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.Rename;
                _ftpRequest.RenameTo = sRenameTo;
                _ftpResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                _ftpResponse.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp create folder failed: " + ex.ToString());
                return false;
            }

            return true;
        }
        public bool RenameUSBStorage(string sDestination, string sRenameTo, ref string exceptionInfo)
        {
            try
            {
                string sTotalDestination = _ftpUrl + sDestination;
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalDestination);
                _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.Rename;
                _ftpRequest.RenameTo = sRenameTo;
                _ftpResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                _ftpResponse.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                exceptionInfo = ex.ToString();
                Debug.WriteLine("Ftp create folder failed: " + ex.ToString());
                return false;
            }

            return true;
        }

        public bool ListDirUSBStorage(string sDestination, ref string sGetDir)
        {
            try
            {
                string sTotalDestination = _ftpUrl + sDestination;
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalDestination);
                _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;

                _ftpResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                Stream responseStream = _ftpResponse.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream);
                sGetDir = reader.ReadToEnd();
                reader.Close();
                _ftpResponse.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp create folder failed: " + ex.ToString());
                return false;
            }
            if (sGetDir.Length == 0)
                return false;
            else
                return true;
        }
        public bool ListDirUSBStorage(string sDestination, ref string sGetDir, ref string exceptionInfo)
        {
            try
            {
                string sTotalDestination = _ftpUrl + sDestination;
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalDestination);
                _ftpRequest.Credentials = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.UsePassive = this._passiveMode;
                _ftpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;

                _ftpResponse = (FtpWebResponse)_ftpRequest.GetResponse();
                Stream responseStream = _ftpResponse.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream);
                sGetDir = reader.ReadToEnd();
                reader.Close();
                _ftpResponse.Close();
                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                exceptionInfo = ex.ToString();
                Debug.WriteLine("Ftp create folder failed: " + ex.ToString());
                return false;
            }
            if (sGetDir.Length == 0)
                return false;
            else
                return true;
        }


        public List<string> GetFtpFileList()
        {
            string str = string.Empty;
            List<string> strList = new List<string>();

            try
            {
                //_ftpFilePath = _ftpUrl + "/" + sFilePath + sFileName;
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(_ftpUrl));
                NetworkCredential ftpCredential = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.Credentials = ftpCredential;
                _ftpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                //_ftpRequest.Method = WebRequestMethods.Ftp.li;

                FtpWebResponse response = (FtpWebResponse)_ftpRequest.GetResponse();
                Stream responseStream = response.GetResponseStream();

                StreamReader reader = new StreamReader(responseStream);

                str = reader.ReadLine();
                while (str != null)
                {
                    strList.Add(str);
                    str = reader.ReadLine();
                }

                reader.Close();
                response.Close();

                reader.Dispose();

                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp upload failed: " + ex.ToString());
                throw new Exception("Ftp upload failed: " + ex.ToString());
                //return false;
            }

            return strList;
        }

        public List<string> GetFtpDirList()
        {
            string str = string.Empty;
            List<string> strList = new List<string>();

            try
            {
                //_ftpFilePath = _ftpUrl + "/" + sFilePath + sFileName;
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(_ftpUrl));
                NetworkCredential ftpCredential = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.Credentials = ftpCredential;
                _ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;

                FtpWebResponse response = (FtpWebResponse)_ftpRequest.GetResponse();
                Stream responseStream = response.GetResponseStream();

                StreamReader reader = new StreamReader(responseStream);

                str = reader.ReadLine();
                while (str != null)
                {
                    strList.Add(str);
                    str = reader.ReadLine();
                }

                reader.Close();
                response.Close();

                reader.Dispose();

                _ftpRequest = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp upload failed: " + ex.ToString());
                throw new Exception("Ftp upload failed: " + ex.ToString());
                //return false;
            }

            return strList;
        }

        public bool GetDownloaddFileSize(string sFilePath, string sFileName, ref string str)
        {
            //FtpWebRequest remoteFileLenReq; // 此請求是為了獲取遠端文件長度
            //FtpWebRequest remoteFileReadReq;// 此請求是為了讀取文件
            Stream readStream = null;       // 讀取流
            FileStream writeStream = null;  // 寫本地文件流           

            try
            {
                string sTotalSrcFilePath = _ftpPrefix + sFilePath + sFileName;

                // 下面程式碼目的是取遠端文件長度                
                _ftpRequest = (FtpWebRequest)FtpWebRequest.Create(sTotalSrcFilePath);
                NetworkCredential ftpCredential = new NetworkCredential(_ftpUserID, _ftpUserPWD);
                _ftpRequest.Credentials = ftpCredential;
                _ftpRequest.UseBinary = true;
                _ftpRequest.ContentOffset = 0;
                _ftpRequest.Method = WebRequestMethods.Ftp.GetFileSize;
                FtpWebResponse rsp = (FtpWebResponse)_ftpRequest.GetResponse();
                long totalByte = rsp.ContentLength;
                rsp.Close();

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Ftp Get Download File Size: " + ex.ToString());
                str += "Exception: Ftp Get Download File Size:" + ex.ToString();
                throw new Exception("Ftp Get Download File Size: " + ex.ToString());
                //return false;
            }
            finally
            {
                if (readStream != null)
                {
                    readStream.Close();
                }
                if (writeStream != null)
                {
                    writeStream.Close();
                }
            }

            return true;
        }





    }
}


/*

 * https://tw.saowen.com/a/82c55576b36d8124387d300eaec4221d0d107f961fd7ebf3e3d8105e462b55f2
 * http://www.aspphp.online/bianchen/dnet/cxiapu/cxprm/201701/186056.html
 
 
 
         //#region 構造函數
        ///// 

        ///// 缺省構造函數
        ///// 

        //public CbtFtpClient()
        //{
        //    this.ftpServerIP = "";
        //    this.remoteFilePath = "";
        //    this.ftpUserID = "";
        //    this.ftpPassword = "";
        //    this.ftpServerPort = 21;
        //    this.bConnected = false;
        //}

        ///// 

        ///// 構造函數
        ///// 

        /////FTP服務器IP地址
        /////當前服務器目錄
        /////Ftp 服務器登錄用戶賬號
        /////Ftp 服務器登錄用戶密碼
        /////FTP服務器端口
        //public CbtFtpClient(string ftpServerIP, string remoteFilePath, string ftpUserID, string ftpPassword, int ftpServerPort, bool anonymousAccess = false)
        //{
        //    this.ftpServerIP = ftpServerIP;
        //    this.remoteFilePath = remoteFilePath;
        //    this.ftpUserID = ftpUserID;
        //    this.ftpPassword = ftpPassword;
        //    this.ftpServerPort = ftpServerPort;
        //    //this.Connect();
        //}
        //#endregion


        //#region 登陸字段、屬性
        ///// 

        ///// FTP服務器IP地址
        ///// 

        //private string ftpServerIP;
        //public string FtpServerIP
        //{
        //    get
        //    {
        //        return ftpServerIP;
        //    }
        //    set
        //    {
        //        this.ftpServerIP = value;
        //    }
        //}
        ///// 

        ///// FTP服務器端口
        ///// 

        //private int ftpServerPort;
        //public int FtpServerPort
        //{
        //    get
        //    {
        //        return ftpServerPort;
        //    }
        //    set
        //    {
        //        this.ftpServerPort = value;
        //    }
        //}
        ///// 

        ///// 當前服務器目錄
        ///// 

        //private string remoteFilePath;
        //public string RemoteFilePath
        //{
        //    get
        //    {
        //        return remoteFilePath;
        //    }
        //    set
        //    {
        //        this.remoteFilePath = value;
        //    }
        //}
        ///// 

        ///// Ftp 服務器登錄用戶賬號
        ///// 

        //private string ftpUserID;
        //public string FtpUserID
        //{
        //    set
        //    {
        //        this.ftpUserID = value;
        //    }
        //}
        ///// 

        ///// Ftp 服務器用戶登錄密碼
        ///// 

        //private string ftpPassword;
        //public string FtpPassword
        //{
        //    set
        //    {
        //        this.ftpPassword = value;
        //    }
        //}

        ///// 

        ///// 是否登錄
        ///// 

        //private bool bConnected;
        //public bool Connected
        //{
        //    get
        //    {
        //        return this.bConnected;
        //    }
        //}
        //#endregion
 
 
 
 
 
 
 
 
 
 
 
 
 */