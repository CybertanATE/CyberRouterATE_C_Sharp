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
using System.Net;
using System.Net.NetworkInformation;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        /* Declare delegate prototype */
        //private delegate void SetTextCallBack(string text);
        private delegate void SetTextCallBack(string text, TextBox textbox);
        private delegate void ButtonFunction(bool stat);
        private delegate void SetMessage(string text);
        
        /* Chariot Relative function  */

        /* Run Chariot Console Command*/
        private void RunChariotConsole(string exeFile, string input, string output)
        {
            /* Check if runtst.exe and fmttst.exe exist in the task manager */
            Process[] ps = Process.GetProcesses();
            string pName = string.Empty;

            foreach (Process p in ps)
            {
                pName = p.ProcessName.ToString().ToLower();
                if (pName.IndexOf("runtst") >= 0 || pName.IndexOf("fmttst") >= 0)
                    p.Kill();
            }
            Thread.Sleep(5000);

            string TSToutput = output + ".tst";
            string CSVoutput = output + ".csv";
            string PDFoutput = output + ".pdf";
            

            PDFoutput = PDFoutput.Substring(0, PDFoutput.Length - Path.GetFileName(PDFoutput).Length) + "PDF";

            string runFile = exeFile.Substring(0, (exeFile.Length - Path.GetFileName(exeFile).Length)) + "runtst.exe";
            string fmtFile = exeFile.Substring(0, (exeFile.Length - Path.GetFileName(exeFile).Length)) + "fmttst.exe";

            RunTstExe(runFile, input, TSToutput, "-v");
            FmtTstExe(fmtFile, TSToutput, CSVoutput, "-v");
            FmtTstExe(fmtFile, TSToutput, PDFoutput, "-p");
            File.Delete(TSToutput);
        }

        private void RunChariotConsole(string tstExe,string fmtExe, string input, string output)
        {
            /* Check if runtst.exe and fmttst.exe exist in the task manager */
            Process[] ps = Process.GetProcesses();
            string pName = string.Empty;

            foreach (Process p in ps)
            {
                pName = p.ProcessName.ToString().ToLower();
                if (pName.IndexOf("runtst") >= 0 || pName.IndexOf("fmttst") >= 0)
                    p.Kill();
            }
            Thread.Sleep(5000);

            string TSToutput = output + ".tst";
            string CSVoutput = output + ".csv";
            string PDFoutput = output + ".pdf";

            PDFoutput = PDFoutput.Substring(0, PDFoutput.Length - Path.GetFileName(PDFoutput).Length) + "PDF\\" + Path.GetFileName(PDFoutput);

            //string runFile = exeFile.Substring(0, (exeFile.Length - Path.GetFileName(exeFile).Length)) + "runtst.exe";
            //string fmtFile = exeFile.Substring(0, (exeFile.Length - Path.GetFileName(exeFile).Length)) + "fmttst.exe";

            RunTstExe(tstExe, input, TSToutput, "-v");
            FmtTstExe(fmtExe, TSToutput, CSVoutput, "-v");
            FmtTstExe(fmtExe, TSToutput, PDFoutput, "-p");
            File.Delete(TSToutput);
        }
        
        public void RunTstExe(string exe_File, string input_file, string output_file, string parameter)
        {
            /* Check runtst.exe is not in the task manager */
            /*
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if (p.ProcessName.ToString().IndexOf("runtst") >= 0)
                    p.Kill();
            }
            Thread.Sleep(5000);
            */

            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = "\"" + exe_File + "\"";
            //startInfo.FileName = chariotExeFile.Replace("ixcariot.exe", "runts);
            //startInfo.WorkingDirectory = working_dir;
            //startInfo.Arguments = input_file + " " + output_file + " -v";
            startInfo.Arguments = "\"" + input_file + "\"" + " " + "\"" + output_file + "\"" + " " + parameter;

            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            process.StartInfo = startInfo;
            process.Start();

            // Read the standard output of the process.
            string output;
            while ((output = process.StandardOutput.ReadLine()) != null || (output = process.StandardError.ReadLine()) != null)
            {
                Invoke(new SetTextCallBack(SetText), new object[] { output });
                //txt_Information_Content.AppendText(output);
                //txt_Information_Content.AppendText(Environment.NewLine);
            }

            process.WaitForExit();
            process.Close();
        }

        /// <summary>
        /// Function FmtTst
        /// </summary>
        /// <param name="working_dir">The folder of Chariot.exe </param>
        /// <param name="input_file">input file name ended in .tst</param>
        /// <param name="output_file">output file name</param>
        /// <param name="parameter">-v to output a .csv file</param>
        public void FmtTstExe(string exe_File, string input_file, string output_file, string parameter)
        {
            /* Check runtst.exe is not in the task manager */
            /*
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if (p.ProcessName.ToString().IndexOf("fmttst") >= 0)
                    p.Kill();
            }
            Thread.Sleep(5000);
            */


            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;

            //startInfo.WorkingDirectory = working_dir;
            startInfo.FileName = "\"" + exe_File + "\"";
            //startInfo.Arguments = input_file + " " + output_file + " -v";
            startInfo.Arguments = "\"" + input_file + "\"" + " " + "\"" + output_file + "\"" + " " + parameter;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            process.StartInfo = startInfo;
            process.Start();

            // Read the standard output of the process.
            string output;
            while ((output = process.StandardOutput.ReadLine()) != null || (output = process.StandardError.ReadLine()) != null)
            {
                Invoke(new SetTextCallBack(SetText), new object[] { output });
                //txt_Information_Content.AppendText(output);
                //txt_Information_Content.AppendText(Environment.NewLine);
            }
            process.WaitForExit();
            process.Close();
            string str = "Save File: " + output_file;   
        }
        
        private void RunChariot(string working_dir, string input, string output)
        {
            string str = "Chariot is running";
            Invoke(new SetTextCallBack(SetText), new object[] { str });

            string TSToutput = output + ".tst";
            string CSVoutput = output + ".csv";

            RunTST(working_dir, input, TSToutput);
            FmtTst(working_dir, TSToutput, CSVoutput, "-v");
            File.Delete(TSToutput);

            str = "Chariot is Finished";
            Invoke(new SetTextCallBack(SetText), new object[] { str });

        }

        public void RunTST(string working_dir, string input_file, string output_file)
        {
            /* Check runtst.exe is not in the task manager */
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if (p.ProcessName.ToString().IndexOf("runtst") >= 0)
                    p.Kill();
            }
            Thread.Sleep(5000);

            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = working_dir + "\\runtst.exe";
            //startInfo.FileName = chariotExeFile.Replace("ixcariot.exe", "runts);
            startInfo.WorkingDirectory = working_dir;
            //startInfo.Arguments = input_file + " " + output_file + " -v";
            startInfo.Arguments = "\"" + input_file + "\"" + " " + "\"" + output_file + "\"" + " -v";

            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            process.StartInfo = startInfo;
            process.Start();

            // Read the standard output of the process.
            string output;
            while ((output = process.StandardOutput.ReadLine()) != null || (output = process.StandardError.ReadLine()) != null)
            {
                Invoke(new SetTextCallBack(SetText), new object[] { output });
            }

            process.WaitForExit();
            process.Close();
        }

        /// <summary>
        /// Function FmtTst
        /// </summary>
        /// <param name="working_dir">The folder of Chariot.exe </param>
        /// <param name="input_file">input file name ended in .tst</param>
        /// <param name="output_file">output file name</param>
        /// <param name="parameter">-v to output a .csv file</param>
        public void FmtTst(string working_dir, string input_file, string output_file, string parameter)
        {
            /* Check runtst.exe is not in the task manager */
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if (p.ProcessName.ToString().IndexOf("fmttst") >= 0)
                    p.Kill();
            }
            Thread.Sleep(5000);



            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;

            startInfo.WorkingDirectory = working_dir;
            startInfo.FileName = working_dir + "\\fmttst.exe";
            //startInfo.Arguments = input_file + " " + output_file + " -v";
            startInfo.Arguments = "\"" + input_file + "\"" + " " + "\"" + output_file + "\"" + " -v";
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            process.StartInfo = startInfo;
            process.Start();

            // Read the standard output of the process.
            string output;
            while ((output = process.StandardOutput.ReadLine()) != null || (output = process.StandardError.ReadLine()) != null)
            {
                Invoke(new SetTextCallBack(SetText), new object[] { output });
            }

            process.WaitForExit();
            process.Close();

            string str = "Save File: " + output_file;
            Invoke(new SetTextCallBack(SetText), new object[] { str });
        }

        /* delegate call back function */
        private void SetText(string msg)
        {
            this.txtRvRTestInformation.AppendText(msg + Environment.NewLine);
        }

        private void showMessage(string text)
        {
            MessageBox.Show(text);
        }

        private void RunButtonOnOff(bool onoff)
        {
            if (onoff == true)
            {
                btnRvRFunctionTestRun.Enabled = true;
                //btn_wireless_throughput_test_Run.Visible = false;
            }
            else
            {
                btnRvRFunctionTestRun.Enabled = false;
                //btn_wireless_throughput_test_Run.Visible = true;
            }
        }

         /* Run default cgi before any wireless setting 
         * For E8350, run Session_Close=1, mfgTest=1 first */
        private bool RunPreCGI(string modelName, string DUTIPaddr, string DUTusername, string DUTpassword)
        {
            string cgi = string.Empty;            
            switch (modelName)
            {
                case "E8350": /* E8350 */
               {
                    /* Run session.cgi?Close_Session=1 */
                    cgi = GetCGICommand(modelName, "CloseSession", DUTIPaddr);
                    if (!SendCGICommand(cgi, DUTusername, DUTpassword, "200 OK"))
                    {
                        MessageBox.Show("CGI command : Close_Session Failed. Check Setting again!.", "Warning");
                        return false;
                    }
                    Thread.Sleep(1000);


                    /* Run mfgtst.cgi?sys_mfgTest=1 */
                    cgi = GetCGICommand(modelName, "mfgTest1", DUTIPaddr);
                    if (!SendCGICommand(cgi, DUTusername, DUTpassword, "200 OK"))
                    {
                        MessageBox.Show("CGI command : mfgTest Failed. Check Setting again!.", "Warning");
                        return false;
                    }
                    Thread.Sleep(1000);                        
                    break;
                }
                case "EA8500X5": /* EA8500X5 */
                {
                    /* Run mfgtst.cgi?sys_mfgTest=1 */
                    cgi = GetCGICommand(modelName, "mfgTest1", DUTIPaddr);
                    //if (!SendCGICommand(cgi, DUTusername, DUTpassword, "200 OK"))


                    if (!SendCGICommand(cgi, DUTusername, DUTpassword, "wireless ConfigurationPass"))
                    {
                        //MessageBox.Show("CGI command : mfgTest Failed. Check Setting again!.", "Warning");
                        //return false;
                    }
                    Thread.Sleep(1000);
                    break;
                }
                default:
                    return false;                     
            }
            return true;
        }

        public bool SendCGICommand(string complete_cgi_command, string username, string password, string compareText)
        {
            string uri = complete_cgi_command;
            string strHtml = string.Empty;
            try
            {
                WebRequest myRequest = WebRequest.Create(uri);
                myRequest.Timeout = 30000;
                myRequest.Credentials = new NetworkCredential(username, password);
                HttpWebRequest myHttpWebRequest = (HttpWebRequest)myRequest;
                myHttpWebRequest.Timeout = 30000;
                myHttpWebRequest.Method = "GET";
                myHttpWebRequest.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; Windows NT 5.2; Windows NT 6.0; Windows NT 6.1; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727; MS-RTC LM 8; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022; .NET CLR 3.0.4506.2152; .NET CLR 3.5.30729; .NET CLR 4.0C; .NET CLR 4.0E)";
                using (WebResponse myWebResponse = myHttpWebRequest.GetResponse())
                {
                    using (Stream myStream = myWebResponse.GetResponseStream())
                    {
                        using (StreamReader myReader = new StreamReader(myStream))
                        {
                            strHtml = myReader.ReadToEnd();
                        }
                    }
                    myWebResponse.Close();
                }
                                
                if (strHtml.ToLower().IndexOf(compareText.ToLower()) != -1) return true;
                //if (strHtml.IndexOf("wireless ConfigurationPass") != -1) 
               //     return true;
                else 
                    return false;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
        }

        private string GetCGICommand(string modelName, string command, string DutIP)
        {
            string cgi = string.Empty;
            switch (modelName)
            {
                case "E8350":
                    if (command == "CloseSession")
                    {
                        //cgi = "http://DUTIPaddr/session.cgi?Close_Session=1";
                        cgi = E8350.CloseSession;
                    }

                    if (command == "mfgTest1")
                    {
                        cgi = E8350.mfgTest1;
                        //cgi = "http://DUTIPaddr/mfgtst.cgi?sys_mfgTest=1";
                    }
                    
                    cgi = cgi.Replace("DUTIPaddr", DutIP) ;
                    
                    break;
                case "EA8500X5":
                    if (command == "mfgTest1")
                    {
                        cgi = EA8500.mfgTest1;                        
                    }

                    cgi = cgi.Replace("DUTIPaddr", DutIP);

                    break;
                default:
                    break;
            }
            
            return cgi;
        }

        private bool PingClient(string hostname, int timeout)
        {
            Ping pClient = new Ping();
            string ip_addr = hostname;
            try
            {
                PingReply pReply = pClient.Send(ip_addr, timeout);

                if (pReply.Status == IPStatus.Success)
                    return true;
                else
                    return false;
            }
            catch
            {
                /* Do nothing */
                return false;                
            }
        }

        private bool CheckIPValid(string ip)
        {
            IPAddress ipaddr;
            if (IPAddress.TryParse(ip, out ipaddr))
            { //valid ip
                return true;
            }
            else
            { //invalid ip
                return false;
            }
        }

        private void RunChariotConsole(string tstExe, string fmtExe, string input, string output, TextBox textbox)
        {
            /* Check if runtst.exe and fmttst.exe exist in the task manager */
            Process[] ps = Process.GetProcesses();
            string pName = string.Empty;

            foreach (Process p in ps)
            {
                pName = p.ProcessName.ToString().ToLower();
                if (pName.IndexOf("runtst") >= 0 || pName.IndexOf("fmttst") >= 0)
                    p.Kill();
            }
            Thread.Sleep(5000);

            string TSToutput = output + ".tst";
            string CSVoutput = output + ".csv";
            string PDFoutput = output + ".pdf";

            PDFoutput = PDFoutput.Substring(0, PDFoutput.Length - Path.GetFileName(PDFoutput).Length) + "PDF\\" + Path.GetFileName(PDFoutput);

            //string runFile = exeFile.Substring(0, (exeFile.Length - Path.GetFileName(exeFile).Length)) + "runtst.exe";
            //string fmtFile = exeFile.Substring(0, (exeFile.Length - Path.GetFileName(exeFile).Length)) + "fmttst.exe";

            RunTstExeT(tstExe, input, TSToutput, "-v", textbox);
            FmtTstExeT(fmtExe, TSToutput, CSVoutput, "-v", textbox);
            FmtTstExeT(fmtExe, TSToutput, PDFoutput, "-p", textbox);
            File.Delete(TSToutput);
        }

        public void RunTstExeT(string exe_File, string input_file, string output_file, string parameter, TextBox textbox)
        {
            /* Check runtst.exe is not in the task manager */
            /*
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if (p.ProcessName.ToString().IndexOf("runtst") >= 0)
                    p.Kill();
            }
            Thread.Sleep(5000);
            */

            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = "\"" + exe_File + "\"";
            //startInfo.FileName = chariotExeFile.Replace("ixcariot.exe", "runts);
            //startInfo.WorkingDirectory = working_dir;
            //startInfo.Arguments = input_file + " " + output_file + " -v";
            startInfo.Arguments = "\"" + input_file + "\"" + " " + "\"" + output_file + "\"" + " " + parameter;

            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            process.StartInfo = startInfo;
            process.Start();

            // Read the standard output of the process.
            string output;
            while ((output = process.StandardOutput.ReadLine()) != null || (output = process.StandardError.ReadLine()) != null)
            {
                Invoke(new SetTextCallBackT(SetText), new object[] { output, textbox });
                //txt_Information_Content.AppendText(output);
                //txt_Information_Content.AppendText(Environment.NewLine);
            }

            process.WaitForExit();
            process.Close();
        }

        /// <summary>
        /// Function FmtTst
        /// </summary>
        /// <param name="working_dir">The folder of Chariot.exe </param>
        /// <param name="input_file">input file name ended in .tst</param>
        /// <param name="output_file">output file name</param>
        /// <param name="parameter">-v to output a .csv file</param>
        public void FmtTstExeT(string exe_File, string input_file, string output_file, string parameter, TextBox textbox)
        {
            /* Check runtst.exe is not in the task manager */
            /*
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if (p.ProcessName.ToString().IndexOf("fmttst") >= 0)
                    p.Kill();
            }
            Thread.Sleep(5000);
            */


            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;

            //startInfo.WorkingDirectory = working_dir;
            startInfo.FileName = "\"" + exe_File + "\"";
            //startInfo.Arguments = input_file + " " + output_file + " -v";
            startInfo.Arguments = "\"" + input_file + "\"" + " " + "\"" + output_file + "\"" + " " + parameter;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            process.StartInfo = startInfo;
            process.Start();

            // Read the standard output of the process.
            string output;
            while ((output = process.StandardOutput.ReadLine()) != null || (output = process.StandardError.ReadLine()) != null)
            {
                Invoke(new SetTextCallBackT(SetText), new object[] { output, textbox });
                //txt_Information_Content.AppendText(output);
                //txt_Information_Content.AppendText(Environment.NewLine);
            }
            process.WaitForExit();
            process.Close();
            string str = "Save File: " + output_file;
        }
        
        


    }
}
