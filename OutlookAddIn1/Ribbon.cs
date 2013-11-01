using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace OutlookAddIn1
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookAddIn1.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

        }

        public void ClickLottie(Office.IRibbonControl control)
        {
            AddWeekReminder(control);
        }

        // This should only be for debugging. Would put a 'Conditional' over it, but can't do that in the XML,
        // so it's just going to be a matter of commenting out the minute option in Ribbon.xml <!-- ... -->
        public void AddMinuteReminder(Office.IRibbonControl control)
        {
            AddReminder(new TimeSpan(0, 0, 1, 0), "one minute"); 

        }

        public void AddDayReminder(Office.IRibbonControl control)
        {
            AddReminder(new TimeSpan(1, 0, 0, 0), "one day"); 
        }

        public void AddThreeDayReminder(Office.IRibbonControl control)
        {
            AddReminder(new TimeSpan(3, 0, 0, 0), "three days"); 
        }

        public void AddWeekReminder(Office.IRibbonControl control)
        {
            AddReminder(new TimeSpan(7, 0, 0, 0), "one week");
        }

        private void AddReminder(TimeSpan TimeToNag, String HowLongText)
        {

            Outlook.MailItem msg = GetCurrentMailItem();
            // TODO: Should do something intelligent to add this to the message rather than simple append (and replace the string appropriately if the button is clicked more than once)
            msg.Body += "\n\n I would appreciate a reply within " + HowLongText + ". Thanks!   -AutoLottie";


            Outlook.UserProperty lottieExpiryProperty = msg.UserProperties.Add(AutoLottieAddIn.LottieDurationProperty, Outlook.OlUserPropertyType.olText);
            lottieExpiryProperty.Value = TimeToNag.ToString();

            // actually want a boolean, but I defy you to figure out how to use an olYesNo type.
            Outlook.UserProperty lottieStartThreadProperty = msg.UserProperties.Add(AutoLottieAddIn.LottieNoCancelProperty, Outlook.OlUserPropertyType.olText);
            lottieStartThreadProperty.Value = "True";
 
        }


        private Outlook.MailItem GetCurrentMailItem()
        {
            var outlook = new Outlook.Application();
            var inspector = outlook.ActiveInspector();
            return inspector.CurrentItem as Outlook.MailItem;
        }

 

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
