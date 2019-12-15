using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;

namespace CyberRouterATE
{
    partial class AboutCyberRouterATE : Form
    {
        public AboutCyberRouterATE()
        {
            //InitializeComponent();
            //this.Text = String.Format("About {0}", AssemblyTitle);
            //this.labelProductName.Text = AssemblyProduct;
            //this.labelVersion.Text = String.Format("Version {0}", AssemblyVersion);
            //this.labelCopyright.Text = AssemblyCopyright;
            //this.labelCompanyName.Text = AssemblyCompany;
            //this.textBoxDescription.Text = AssemblyDescription;
            //InitializeComponent();
            //this.Text = String.Format("About {0}", AssemblyTitle);
            //this.labelProductName.Text = AssemblyProduct;
            //this.labelVersion.Text = String.Format("Version {0}", AssemblyVersion);
            //this.labelCopyright.Text = AssemblyCopyright;
            //this.labelCompanyName.Text = AssemblyCompany;
            //this.textBoxDescription.Text = AssemblyDescription;

            InitializeComponent();

            //  Initialize the AboutBox to display the product information from the assembly information.
            //  Change assembly information settings for your application through either:
            //  - Project->Properties->Application->Assembly Information
            //  - AssemblyInfo.cs

            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);

            //this.Text = String.Format("About {0}", AssemblyTitle);
            //this.labelProductName.Text = AssemblyProduct;
            this.labelVersion.Text = String.Format("Version: {0}", fvi.FileVersion);
            //this.labelCopyright.Text = AssemblyCopyright;
            //this.labelCompanyName.Text = AssemblyCompany;
            //this.textBoxDescription.Text = AssemblyDescription;
            this.textBoxDescription.Text = @"Description: This program CyberRouterATE is CyberTan automated test application for Router series product. The CyberRouterATE includes test items:"
            + Environment.NewLine + "------------------------------------------------------"
            + Environment.NewLine + " 1. RvR Test"
            + Environment.NewLine + " 2. Power On/Off Test"
            + Environment.NewLine + " 3. RvR-Turn Table Test"
            + Environment.NewLine + " 4. Interoperability Test"
            + Environment.NewLine + " 5. Throughput Test"
            + Environment.NewLine + " 6. Web GUI Test - FW Upgrade/Downgrade Test"
            + Environment.NewLine + " 7. USB Storage Test"
            + Environment.NewLine + " 8. GUI Test";
            //+ Environment.NewLine + " 3. Turn RvR Test";
            //+ Environment.NewLine + " 4. ISB8K NorDig Task 2-4"
            //+ Environment.NewLine + " 5. ISB8K NorDig Task 2-5"
            //+ Environment.NewLine + " 6. ISB8K NorDig Task 2-13"
            //+ Environment.NewLine + " 7. ISB8K NorDig Task 2-14"
            //+ Environment.NewLine + " 8. ISB8K NorDig Task 2-16"
            //+ Environment.NewLine + " 9. ISB8K NorDig Task 2-17"
            //+ Environment.NewLine + "10. ISB8K NorDig Task 2-18"
            //+ Environment.NewLine + "11. CableModemDownstream"
            //+ Environment.NewLine + "12. CableModemUpstream"
            //+ Environment.NewLine + "13. HDDTA RFI"
            //+ Environment.NewLine + "14. EPC3028/EPC4911 Calibration";
        }

        #region Assembly Attribute Accessors

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion
    }
}
