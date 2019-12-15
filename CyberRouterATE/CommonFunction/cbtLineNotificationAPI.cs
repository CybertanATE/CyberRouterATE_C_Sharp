using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace NS_cbtLineNotificationApi
{
    class CbtLineNotificationApi
    {
        private string _authorization = null;
        private string _apiAddress = @"https://notify-api.line.me/api/notify";
        public static Dictionary<int, string> dic_LineAuth = new Dictionary<int, string>() {
            {1, "dyAWp23TMERyPjEhjFjvZnh2vDKG3Tt56w21N6DL5OB"},
            {2, "sYiVOsbyIZ1bkaclJd4lfk66yE9XDN7umVnmz1LZ1U2"},
            {3, "LQIUps8IjSuzFUAZqkwEMaNIYhrrJkQOkMZo9vgYR1X"},
            {4, "iTM2MuFlXGB1nhSEPNr9ZngpQVOyDd0bI1VbGKAddux"},
            {5, "PgJKnNVDOTuFpDVy8P1a9JnLvpNPB5o2xBUZVbWb2Ns"},
            {6, "UBIqK0aIBAD76Quy2jQld6AncE0tK9trouf9Gv3nipw"} };

        public CbtLineNotificationApi(string authorization)
        {
            this._authorization = authorization;
        }

        public bool postMessage(string message)
        {
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = @"C:\Windows\System32" + "\\cmd.exe";
            startInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;

            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            process.StartInfo = startInfo;
            process.Start();
            process.StandardInput.WriteLine("cd " + System.Windows.Forms.Application.StartupPath + @"\Archives\curl_7630win64mingw\bin");
            process.StandardInput.WriteLine(@"curl -X POST -H ""Authorization:Bearer " + _authorization + @""" -F ""message=" + message + @""" " + _apiAddress);
            process.StandardInput.WriteLine("exit");
            string tempOutput = "";
            string output = "";
            while ((tempOutput = process.StandardOutput.ReadLine()) != null || (tempOutput = process.StandardError.ReadLine()) != null)
            {
                output += tempOutput + "\n";
            }
            string expectedOutput = @"""status"":200,""message"":""ok""";
            if (output.Contains(expectedOutput) == false)
            {
                process.WaitForExit();
                process.Close();
                return false;
            }
            process.WaitForExit();
            process.Close();
            return true;
        }

        public bool postMessageAndSticker(string message, int stickerPackageId, int stickerId)
        {
            //stickerPackageId = 2, stickerId = 144  熊大歡呼
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = @"C:\Windows\System32" + "\\cmd.exe";
            startInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;

            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            process.StartInfo = startInfo;
            process.Start();
            process.StandardInput.WriteLine("cd " + System.Windows.Forms.Application.StartupPath + @"\Archives\curl_7630win64mingw\bin");
            process.StandardInput.WriteLine(@"curl -X POST -H ""Authorization:Bearer " + _authorization + @""" -F ""message=" + message + @""" -F ""stickerPackageId=" + stickerPackageId.ToString() + @""" -F ""stickerId=" + stickerId.ToString() + @""" " + _apiAddress);
            process.StandardInput.WriteLine("exit");
            string tempOutput = "";
            string output = "";
            while ((tempOutput = process.StandardOutput.ReadLine()) != null || (tempOutput = process.StandardError.ReadLine()) != null)
            {
                output += tempOutput + "\n";
            }
            string expectedOutput = @"""status"":200,""message"":""ok""";
            if (output.Contains(expectedOutput) == false)
            {
                process.WaitForExit();
                process.Close();
                return false;
            }
            process.WaitForExit();
            process.Close();
            return true;
        }

        public bool postMessageAndPicture(string message, string picturePath)
        {
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = @"C:\Windows\System32" + "\\cmd.exe";
            startInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;

            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            process.StartInfo = startInfo;
            process.Start();
            process.StandardInput.WriteLine("cd " + System.Windows.Forms.Application.StartupPath + @"\Archives\curl_7630win64mingw\bin");
            process.StandardInput.WriteLine(@"curl -X POST -H ""Authorization:Bearer " + _authorization + @""" -F ""message=" + message + @""" -F ""imageFile=@" + picturePath + @""" " + _apiAddress);
            process.StandardInput.WriteLine("exit");
            // 語法 curl -X POST https://notify-api.line.me/api/notify -H "Authorization:Bearer [Accesstoken]" -F "message=[必須附帶訊息]" -F “imageFile=@[路路徑＋檔名].jpg”
            string tempOutput = "";
            string output = "";
            while ((tempOutput = process.StandardOutput.ReadLine()) != null || (tempOutput = process.StandardError.ReadLine()) != null)
            {
                output += tempOutput + "\n";
            }
            string expectedOutput = @"""status"":200,""message"":""ok""";
            if (output.Contains(expectedOutput) == false)
            {
                process.WaitForExit();
                process.Close();
                return false;
            }
            process.WaitForExit();
            process.Close();
            return true;
        }

        public static void createDic()
        {
            dic_LineAuth.Add(1, "dyAWp23TMERyPjEhjFjvZnh2vDKG3Tt56w21N6DL5OB");
            dic_LineAuth.Add(2, "sYiVOsbyIZ1bkaclJd4lfk66yE9XDN7umVnmz1LZ1U2");
            dic_LineAuth.Add(3, "LQIUps8IjSuzFUAZqkwEMaNIYhrrJkQOkMZo9vgYR1X");
            dic_LineAuth.Add(4, "iTM2MuFlXGB1nhSEPNr9ZngpQVOyDd0bI1VbGKAddux");
            dic_LineAuth.Add(5, "PgJKnNVDOTuFpDVy8P1a9JnLvpNPB5o2xBUZVbWb2Ns");
            dic_LineAuth.Add(6, "UBIqK0aIBAD76Quy2jQld6AncE0tK9trouf9Gv3nipw");

        }


    }
}
