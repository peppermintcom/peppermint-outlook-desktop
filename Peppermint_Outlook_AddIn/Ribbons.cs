using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Drawing;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;


// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Peppermint_Outlook_AddIn
{
    [ComVisible(true)]
    public class Ribbons : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbons()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Peppermint_Outlook_AddIn.Ribbons.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void btnRecordMessage_Click(Office.IRibbonControl control)
        {
            Outlook.MailItem mi = ThisAddIn.theCurrentMailItem.ReplyAll();
            mi.Display();
            ThisAddIn.RecordAudioAndAttach("Read");
        }

        public void btnRecordMessageNewMail_Click(Office.IRibbonControl control)
        {
            ThisAddIn.theCurrentMailItem.Save();
            ThisAddIn.RecordAudioAndAttach("Create");
        }

        public void btnSendViaPeppermint_Click(Office.IRibbonControl control)
        {
            try
            {
                if (ThisAddIn.outlookApp.ActiveExplorer().Selection == null);
            }
            catch (Exception)
            {
                RecordAndAttachAudio();
                return;
            }

            if (ThisAddIn.outlookApp.ActiveExplorer().Selection.Count <= 0)
            {
                RecordAndAttachAudio();
                return;
            }

            // If an email is selected Reply to that email via Peppermint, if none is selected then Start and audio recording, 
            // else if more then 1 -mail is selected, prompt the end-user to select a single e-mail 
            if (ThisAddIn.outlookApp.ActiveExplorer().Selection.Count > 1)
            {
                MessageBox.Show("Please select a single e-mail to respond to via Peppermint", "More than one e-mail selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (ThisAddIn.outlookApp.ActiveExplorer().Selection.Count == 1)
                if (ThisAddIn.outlookApp.ActiveExplorer().Selection[1] is Outlook.MailItem)
                {
                    ThisAddIn.theCurrentMailItem = ThisAddIn.outlookApp.ActiveExplorer().Selection[1] as Outlook.MailItem;

                    Outlook.MailItem mi = ThisAddIn.theCurrentMailItem.ReplyAll();
                    mi.Display();
                    ThisAddIn.RecordAudioAndAttach("Read");
                }
        }

        private void RecordAndAttachAudio()
        {
            if (ThisAddIn.RecordAudioAndAttach("Explorer") == DialogResult.OK)
            {
                // Create a new email only of the audio is to be attached/sent
                ThisAddIn.theCurrentMailItem = ThisAddIn.outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                ThisAddIn.theCurrentMailItem.Display();

                ThisAddIn.theCurrentMailItem.Subject = "I sent you a voicemail message";

                ThisAddIn.theCurrentMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                ThisAddIn.bPeppermintMessageInserted = true;

                // Attach audio recording file
                if ((ThisAddIn.theCurrentMailItem != null) && (File.Exists(ThisAddIn.AttachmentFilePath)))
                    ThisAddIn.theCurrentMailItem.Attachments.Add(ThisAddIn.AttachmentFilePath);

                if (!String.IsNullOrEmpty(ThisAddIn.PEPPERMINT_TRANSCRIBED_AUDIO))
                {
                    ThisAddIn.theCurrentMailItem.HTMLBody = ThisAddIn.PEPPERMINT_TRANSCRIBED_TEXT_HEADER +
                                                            ThisAddIn.PEPPERMINT_TRANSCRIBED_AUDIO +
                                                            ThisAddIn.theCurrentMailItem.HTMLBody;
                }
                ThisAddIn.theCurrentMailItem.HTMLBody = ThisAddIn.PEPPERMINT_NEW_EMAIL_HTML_BODY + ThisAddIn.theCurrentMailItem.HTMLBody;
            }
        }
        public Bitmap btnRecordMessage_getImage(Office.IRibbonControl control)
        {
            return Properties.Resources.Logo;
        }

        public void btnFeedback_Click(Office.IRibbonControl control)
        {
            Outlook.MailItem mi = ThisAddIn.outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

            mi.Subject = "Feedback : Peppermint Outlook AddIn";
            mi.Recipients.Add(ThisAddIn.PEPPERMINT_SUPPORT_EMAIL);

            string strOSBitness = string.Empty;
            string strOfficeBitness = string.Empty;

            if (Environment.Is64BitOperatingSystem == true)
            {
                strOSBitness = "64-bit";
            }
            else
            {
                strOSBitness = "32-bit";
            }

            if (Environment.Is64BitProcess == true)
            {
                strOfficeBitness = "64-bit";
            }
            else
            {
                strOfficeBitness = "32-bit";
            }

            string AddinVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();

            mi.Body += "\r\n\r\n\r\n" + "O.S. version : " + ThisAddIn.GetOSFriendlyName() + "\t" + strOSBitness +
                                    "\r\nOutlook version : " + ThisAddIn.outlookApp.Version + "\t" + strOfficeBitness +
                                     "\r\nPeppermint AddIn version : " + AddinVersion;

            mi.Display();
        }

        public void btnAbout_Click(Office.IRibbonControl control)
        {
            frmAbout myAboutFrom = new frmAbout();
            myAboutFrom.ShowDialog();
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
