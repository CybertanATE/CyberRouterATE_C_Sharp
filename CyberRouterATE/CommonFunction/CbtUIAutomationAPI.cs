///---------------------------------------------------------------------------------------
///  Created by CyberTan.
///  File           : CbtUIAutomationAPI.cs
///  Update         : 2015-11-09  
///  Version        : 
///  Description    : Encapsulate Microsoft UIAutomation class which is provided to CyberTan 
///                   internal ATE development.
///                   Refer to: https://msdn.microsoft.com/en-us/library/ms747327(v=vs.100).aspx
///                   Note: If you're finding an error about Rect class in your code, here's what you
///                   need to do:
///                   1. Add WindowsBase.dll reference file
///                   2. Browse the path of c:\Program Files\Reference Assemblies\Microsoft\Framework\v3.0\
///                      to find WindowsBase.dll file, but I CAN NOT make sure which path you have in your PC.
///  Modified       : 2015-11-09 Initial version
///---------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Automation;
using Automation = System.Windows.Automation;

namespace CyberATE.CommonAPI.CbtUIAutomationAPI
{
    class CbtUI
    {
        /* Get element by ClassName property */
        public AutomationElement GetElementByClassName(AutomationElement parentElement, string classname, ControlType ctrlType)
        {
            if (parentElement == null || ctrlType == null)
            {
                throw new InvalidOperationException("Could not find the element!");                
            }

            AndCondition andCondition = new AndCondition(new PropertyCondition(AutomationElement.ClassNameProperty, classname), new PropertyCondition(AutomationElement.ControlTypeProperty, ctrlType));
            return parentElement.FindFirst(TreeScope.Descendants | TreeScope.Element, andCondition);
        }

        /* Get element by Name property */
        public AutomationElement GetElementByName(AutomationElement parentElement, string name, ControlType ctrlType)
        {
            if (parentElement == null || ctrlType == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }

            AndCondition andCondition = new AndCondition(new PropertyCondition(AutomationElement.NameProperty, name), new PropertyCondition(AutomationElement.ControlTypeProperty, ctrlType));
            return parentElement.FindFirst(TreeScope.Descendants | TreeScope.Element, andCondition);
        }

        /* Get element by AutomationId property */
        public AutomationElement GetElementByAutomationId(AutomationElement parentElement, string id, ControlType ctrlType)
        {
            if (parentElement == null || ctrlType == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }

            AndCondition andCondition = new AndCondition(new PropertyCondition(AutomationElement.AutomationIdProperty, id), new PropertyCondition(AutomationElement.ControlTypeProperty, ctrlType));
            return parentElement.FindFirst(TreeScope.Descendants | TreeScope.Element, andCondition);
        }

        /* Get element by AutomationId property */
        public AutomationElement GetElementByProcessId(AutomationElement parentElement, int id, ControlType ctrlType)
        {
            if (parentElement == null || ctrlType == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }

            AndCondition andCondition = new AndCondition(new PropertyCondition(AutomationElement.ProcessIdProperty, id), new PropertyCondition(AutomationElement.ControlTypeProperty, ctrlType));
            return parentElement.FindFirst(TreeScope.Descendants | TreeScope.Element, andCondition);
        }

        /* Get element by AutomationId property */
        public AutomationElement GetElementByBoundingRectangle(AutomationElement parentElement, Rect rect)
        {
            if (parentElement == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }

            PropertyCondition bondingProperty = new PropertyCondition(AutomationElement.BoundingRectangleProperty, rect);
            return parentElement.FindFirst(TreeScope.Descendants | TreeScope.Element, bondingProperty);
        }

        /* Send click button event by AutomationId property */
        public void ClickButtonByAutomationId(AutomationElement parentElement, string id)
        {
            if (parentElement == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }

            AutomationElement element = parentElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
            InvokePattern invokePattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();
        }

        /* Send click button event by CtrlType property */
        public void ClickButtonByCtrlType(AutomationElement parentElement, ControlType ctrlType)
        {
            if (parentElement == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }

            AutomationElement element = parentElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ctrlType));
            InvokePattern invokePattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();
        }

        /* Send click button event by CtrlType property */
        public void ClickButtonByAutomationElement(AutomationElement parentElement)
        {
            if (parentElement == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }

            //parentElement.SetFocus();
            //AutomationElement element = parentElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ctrlType));
            InvokePattern invokePattern = parentElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();
        }

        /* Expand combo box by Automation id */
        public AutomationElement ExpandComboBoxByAutomationId(AutomationElement parentElement, string id)
        {
            if (parentElement == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }

            AutomationElement element = parentElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
            ExpandCollapsePattern expandCollapsePattern = element.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
            expandCollapsePattern.Expand();
            return element;
        }

        /* Get the textbox and put something */
        public void SendTextByAutomationId(AutomationElement parentElement, string id, string text)
        {
            if (parentElement == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }
            AutomationElement element = parentElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
            ValuePattern valuePattern = element.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            valuePattern.SetValue(text);            
        }

        /* Seletc item on ListBox */
        public void SelectListBoxItem(AutomationElement parentElement)
        {
            if (parentElement == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }
            SelectionItemPattern ItemToSelect = parentElement.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
            ItemToSelect.Select();
        }

        /* Get the text by Automation id */
        public string GetTextElementByAutomationId(AutomationElement parentElement, string id)
        {
            AutomationElement txtElement = parentElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
            if (parentElement == null || txtElement == null)
            {
                throw new Exception("UI element is null!");
            }
            return txtElement.Current.Name.ToString();
        }

        /* Get the text by Automation id */
        public string GetTextElementByAutomationElement(AutomationElement parentElement)
        {
            //AutomationElement txtElement = parentElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
            if (parentElement == null)
            {
                throw new Exception("UI element is null!");
            }
            return parentElement.Current.Name.ToString();
        }

        public void SetTextByAutomationElement(AutomationElement parentElement, string text)
        {
            if (parentElement == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }
            //parentElement.SetFocus();
            //AutomationElement element = parentElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
            InvokePattern invokePattern = parentElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();
            ValuePattern valuePattern = parentElement.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            valuePattern.SetValue(text);
        }


        ///* Get the text content by Automation id */
        //public string GetTextContentElementByAutomationId(AutomationElement parentElement, string id)
        //{
        //    AutomationElement txtElement = parentElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
        //    if (parentElement == null || txtElement == null)
        //    {
        //        throw new Exception("UI element is null!");
        //    }
        //    return txtElement.Current.Name.ToString();
        //}

        /* Expand Menu Ite, by Automation id */
        public void ExpandMenuItem(AutomationElement parentElement)
        {
            if (parentElement == null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }

            ExpandCollapsePattern expandCollapsePattern = GetExpandCollapsePattern(parentElement);

            if (expandCollapsePattern == null)
            {
                return;
            }

            if (expandCollapsePattern.Current.ExpandCollapseState ==
                ExpandCollapseState.LeafNode)
            {
                return;
            }

            try
            {
                if (expandCollapsePattern.Current.ExpandCollapseState == ExpandCollapseState.Expanded)
                {
                    // Collapse the menu item.
                    expandCollapsePattern.Collapse();
                }
                else if (expandCollapsePattern.Current.ExpandCollapseState == ExpandCollapseState.Collapsed ||
                    expandCollapsePattern.Current.ExpandCollapseState == ExpandCollapseState.PartiallyExpanded)
                {
                    // Expand the menu item.
                    expandCollapsePattern.Expand();
                }
            }
            // Control is not enabled
            catch (ElementNotEnabledException)
            {
                // TO DO: error handling.
            }
            // Control is unable to perform operation.
            catch (InvalidOperationException)
            {
                // TO DO: error handling.
            }

            //ExpandCollapsePattern pattern = parentElement.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;

            //SelectionItemPattern ItemToSelect = parentElement.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
            //ItemToSelect.Select();
        }

        private ExpandCollapsePattern GetExpandCollapsePattern(AutomationElement targetControl)
        {
            ExpandCollapsePattern expandCollapsePattern = null;

            try
            {
                expandCollapsePattern =
                    targetControl.GetCurrentPattern(
                    ExpandCollapsePattern.Pattern)
                    as ExpandCollapsePattern;
            }
            // Object doesn't support the ExpandCollapsePattern control pattern.
            catch (InvalidOperationException)
            {
                return null;
            }

            return expandCollapsePattern;
        }

        /* Click MenuIte */
        /* Expand combo box by Automation id */
        public void ClickMenuItembyAutomationId(AutomationElement barElemnet, AutomationElement parentElement, string id, ControlType ctrlType)
        {
            if (parentElement == null || barElemnet== null)
            {
                throw new InvalidOperationException("Could not find the element!");
            }

            try
            {
                ExpandCollapsePattern expandCollapsePattern = barElemnet.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
                expandCollapsePattern.Expand();

                AndCondition andCondition = new AndCondition(new PropertyCondition(AutomationElement.AutomationIdProperty, id), new PropertyCondition(AutomationElement.ControlTypeProperty, ctrlType));
                AutomationElement element = parentElement.FindFirst(TreeScope.Descendants | TreeScope.Element, andCondition);

                if (element == null) return;
                else
                {
                    element.SetFocus();
                    InvokePattern invokePattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    invokePattern.Invoke();
                }
            }
            catch (InvalidOperationException)
            {
                return;
            }

            
        }

    }
}
