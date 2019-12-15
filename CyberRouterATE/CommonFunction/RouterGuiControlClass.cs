using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WatiN.Core;
using System.Diagnostics;

namespace RouterControlClass
{
    class RouterGuiControl
    {
        private WatiN.Core.IE _ie = null;
        IntPtr hwnd;
        static private string DutConfigurationType = null;
        static private string DutConfigurationFile = null;

        private const bool ATE_FAIL = false;
        private const bool ATE_PASS = true;

        public RouterGuiControl(string host)
        {
            try
            {
                //First we open a new Internet Explorer browser, making sure it will be visible...                
                _ie = new IE() { Visible = true };
                _ie.GoTo(host);
                _ie.WaitForComplete();
                hwnd = _ie.hWnd;
            }
            catch (Exception ex)
            {                
                Debug.WriteLine(ex.ToString());
                Debug.WriteLine("IE Open Failed : A00");
                _ie = null;
            }                
        }

        public RouterGuiControl(string dutConfigurationType, string dutConfigurationFile)
        {
            try
            {
                if (DutConfigurationType == null)
                    DutConfigurationType = dutConfigurationType;
                if (DutConfigurationFile == null)
                    DutConfigurationFile = dutConfigurationFile;                
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                Debug.WriteLine("Configuration Data Setup Failed : A00");
                _ie = null;
            }
        }

         ~RouterGuiControl()
        {
            Debug.WriteLine("Discontruction");
        }

        public bool Dispose()
        {
            if (_ie != null)
                _ie.Close();
            return ATE_PASS;
        }

        public void Close()
        {
            if (_ie != null) _ie.Close();
        }

        public bool LoginByTypeName(string host, string Type1Name, string username, string Type2Name, string password, string SumitType, string SubmitName, string action)
        {
            try
            {
                _ie.GoTo(host);
                _ie.WaitForComplete();

                TextField searchUsername = _ie.TextField(Find.ByName(Type1Name));
                TextField searchPassword = _ie.TextField(Find.ByName(Type2Name));

                Button btn = null;

                switch (SumitType)
                {
                    case "Button":
                        btn = _ie.Button(Find.ById(SubmitName));                        
                        break;
                    case "Value":
                        btn = _ie.Button(Find.ByValue(SubmitName));
                        break;
                    default:
                        Debug.WriteLine("Unkonwn Type: A01");
                        return ATE_FAIL;
                }
                
                if (searchUsername == null || searchPassword == null || btn == null)
                {
                    Debug.WriteLine("Login type search failed. : A01");
                    return false;
                }

                searchUsername.Value = username;
                searchPassword.Value = password;
                _ie.NativeDocument.Body.SetFocus();
                btn.Click();
                _ie.WaitForComplete();
                string strHtml = _ie.Html;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Login failed : A01");
                return ATE_FAIL;
            }          

            return ATE_PASS;
        }

        public bool LoginByTypeId(string host, string IdName, string password, string SumitType, string SubmitName, string action)
        {
            try
            {
                AttachHostByUrl(host);

                TextField searchPassword = _ie.TextField(Find.ById(IdName));

                Button btn = null;

                switch (SumitType)
                {
                    case "Button":
                        btn = _ie.Button(Find.ById(SubmitName));
                        break;
                    case "Value":
                        btn = _ie.Button(Find.ByValue(SubmitName));
                        break;
                    case "Id":                        
                        btn = _ie.Button(Find.ById(SubmitName));
                        break;
                    default:
                        Debug.WriteLine("Unkonwn Type: A01");
                        return ATE_FAIL;
                }

                if (searchPassword == null || btn == null)
                {
                    Debug.WriteLine("Login type search failed. : A01");
                    return false;
                }
                                
                searchPassword.Value = password;
                _ie.NativeDocument.Body.SetFocus();
                btn.Click();
                _ie.WaitForComplete();
                string strHtml = _ie.Html;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Login failed : A01");
                return ATE_FAIL;
            }

            return ATE_PASS;
        }

        public bool LogoutByTypeName(string host, string Type1Name, string username, string Type2Name, string password, string SumitType, string SubmitName, string action)
        {

            return ATE_PASS;
        }

        public bool LogoutByTypeID(string host, string Type1Name, string username, string Type2Name, string password, string SumitType, string SubmitName, string action)
        {

            return ATE_PASS;
        }

        public bool ButtonClick(string host, string Type1Name, string username, string Type2Name, string password, string SumitType, string SubmitName, string action)
        {

            return ATE_PASS;
        }

        public TextField GetTextFieldbyName(string host, IE ie, string name)
        {
            TextField searchField = null;
            try
            {
                AttachHostByUrl(host);
                searchField = _ie.TextField(Find.ByName(name));                
            }
            catch(Exception ex)
            {
                Debug.WriteLine("Get Text Field failed: " + ex.ToString());
            }

            return searchField;
        }

        public TextField GetTextFieldbyID(string host, IE ie, string id)
        {
            TextField searchField = null;
            try
            {
                AttachHostByUrl(host);
                searchField = _ie.TextField(Find.ById(id));
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Get Text Field failed: "+ ex.ToString());
            }

            return searchField;
        }

        public string GetTextFieldValue(string host, TextField textField)
        {
            string value = null;

            try
            {
                AttachHostByUrl(host);
                value  = textField.Value;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Get Text Field Value failed: " + ex.ToString());
            }
            return value;
        }

        public bool ConfigureTextFieldValue(string host, TextField textField, string value)
        {
            try
            {
                AttachHostByUrl(host);
                //textField.TypeText(value);
                textField.Value = value;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Set Text Field Value failed: " + ex.ToString());
                return ATE_FAIL;
            }

            return ATE_PASS;
        }

        public bool FindTableAllTextByClass(string host, string className, ref string[,] tableData)
        {
            try
            {
                AttachHostByUrl(host);
                Table table = _ie.Table(Find.ByClass(className));
                //table.TableRows.Count
                TableRowCollection trc = table.TableRows;
                int i = 0;
                int j = 0;
                foreach (TableRow tr in trc)
                {
                    TableCellCollection tcc = tr.TableCells;
                    j = 0;
                    foreach (TableCell tc in tcc)
                    {
                        tableData[i, j++] = tc.Text;
                    }
                    i++;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Get Select List All Text failed: " + ex.ToString());
                return ATE_FAIL;
            }
            return ATE_PASS;

        }

        public bool FindTableTextByClassName(string host, string className, int indexRow, int indexColumn, ref string value)
        {
            try
            {
                AttachHostByUrl(host);
                Table table = _ie.Table(Find.ByClass(className));
                TableRowCollection trc = table.TableRows;


                for (int i = 0; i < trc.Count; i++)
                {
                    if (i == indexRow)
                    {
                        TableCellCollection tcc = trc[i].TableCells;
                        value  = tcc[indexColumn].Text;
                        break;                        
                    }                   
                }               
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Get Select List Text failed: " + ex.ToString());
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        public bool FindTableTextByClassName(string host, string className, string searchValue, int indexColumn, ref string value)
        {
            try
            {
                AttachHostByUrl(host);
                Table table = _ie.Table(Find.ByClass(className));
                TableRowCollection trc = table.TableRows;

                foreach (TableRow tr in trc)
                {
                    TableCellCollection tcc = tr.TableCells;
                    if (tcc[0].Text.Trim() == searchValue)
                    {
                        value = tcc[indexColumn].Text;
                        break;
                    }                    
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Get Select List Text failed: " + ex.ToString());
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        public bool Goto(string host)
        {
            try
            {
                _ie.GoTo(host);
                _ie.WaitForComplete();
            }
            catch(Exception ex)
            {
                Debug.WriteLine("Browser host failed: " + ex.ToString());
                return ATE_FAIL;
            }

            return ATE_PASS;
        }

        public bool AttachHost(string host)
        {
            try
            {
                _ie = IE.AttachTo<IE>(Find.ByUrl(host));                    
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Browser host failed: " + ex.ToString());
                return ATE_FAIL;
            }

            return ATE_PASS;
        }
        public bool AttachHostByUrl(string host)
        {
            try
            {
                _ie = IE.AttachTo<IE>(Find.ByUrl(host));                
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Browser host failed: " + ex.ToString());
                return ATE_FAIL;
            }

            return ATE_PASS;
        }

        public bool AttachHostByTitle(string Title)
        {
            try
            {
                _ie = IE.AttachTo<IE>(Find.ByTitle(Title));
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Browser host failed: " + ex.ToString());
                return ATE_FAIL;
            }

            return ATE_PASS;
        }

        public bool AttachHostByTitle(string Title, ref string strHtml)
        {
            try
            {
                _ie = IE.AttachTo<IE>(Find.ByTitle(Title));
                strHtml = _ie.Html;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Browser host failed: " + ex.ToString());
                return ATE_FAIL;
            }

            return ATE_PASS;
        }

        public bool AttachHostByUrl(string host, ref string strHtml)
        {
            try
            {
                _ie = IE.AttachTo<IE>(Find.ByUrl(host));
                strHtml = _ie.Html;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Browser host failed: " + ex.ToString());
                return ATE_FAIL;
            }

            return ATE_PASS;
        }

        public bool ConfigureSelectListValueByName(string host, string name, string value)
        {
            try
            {
                AttachHostByUrl(host);
                // Now we need to find the button to be clicked to launch the search...
                SelectList launchSearch = _ie.SelectList(Find.ByName(name));
                launchSearch.Option(value).Select();
                if (!launchSearch.Exists)
                {
                    Debug.WriteLine("Configure SelectList value by name failed!");
                    return ATE_FAIL;
                }

                _ie.NativeDocument.Body.SetFocus(); // Focusing on the document...
                launchSearch.Click(); // And launching the search...                 
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Configure SelectList value by name failed: " + ex.ToString());
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        public bool ConfigureSelectListValueByNameSelectByValue(string host, string name, string value)
        {
            try
            {
                AttachHostByUrl(host);
                // Now we need to find the button to be clicked to launch the search...
                SelectList launchSearch = _ie.SelectList(Find.ByName(name));
                //launchSearch.Option(value).Select();
                if (!launchSearch.Exists)
                {
                    Debug.WriteLine("Configure SelectList value by name failed!");
                    return ATE_FAIL;
                }

                launchSearch.SelectByValue(value);

                //_ie.NativeDocument.Body.SetFocus(); // Focusing on the document...
                //launchSearch.Click(); // And launching the search...                 
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Configure SelectList value by name failed: " + ex.ToString());
                return ATE_FAIL;
            }
            return ATE_PASS;
        }

        public bool ClickButtonByValue(string host, string value)
        {
            try
            {
                AttachHostByUrl(host);
                Button btn = null;
                btn = _ie.Button(Find.ByValue(value));                

                if (btn == null)
                {
                    Debug.WriteLine("Click Button by Value failed.");
                    return false;
                }
                
                _ie.NativeDocument.Body.SetFocus();
                btn.Click();
                _ie.WaitForComplete();                
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Click Button by Value failed : A01");
                return ATE_FAIL;
            }

            return ATE_PASS;
        }

        public bool Test()
        {
            //SelectListCollection collection = _ie.SelectLists(
            return ATE_PASS;
        }

        public bool checkDutFile()
        {
            if (DutConfigurationFile != null && DutConfigurationType != null)
                return ATE_PASS;
            else
                return ATE_FAIL;
        }

    }
}
