using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;


namespace Peppermint_Outlook_AddIn
{
    public partial class ThisAddIn
    {
        #region fields

        private Outlook.Inspectors _inspectors;
        public static Outlook.MailItem theCurrentMailItem;
        public static Outlook.Application outlookApp;
        public static string AttachmentFilePath;
        public static string PEPPERMINT_NEW_EMAIL_HTML_BODY = "<FONT>Here's my message<BR> Reply via <a href=Peppermint.com>Peppermint.com</a><BR><BR></FONT>";
        public static string PEPPERMINT_REPLY_MAIL_HTML_BODY = "<FONT>I sent you an audio reply with <a href=Peppermint.com>Peppermint.com</a><BR><BR></FONT>";
        public static string PEPPERMINT_WEBSITE = "Peppermint.com";
        public static string PEPPERMINT_SUPPORT_EMAIL = "support@peppermint.com";

        public static string PEPPERMINT_NEW_MAIL_SUBJECT = "I sent you a voicemail message";
        

        public static bool bPeppermintMessageInserted;

        #endregion

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // VSTO Runtime Update to Address Slow Shutdown and “Unknown Publisher” for SHA256 Certificate
            // http://softwareblog.morlok.net/2014/12/03/unknown-publisher-when-installing-clickonce-vsto-outlook-plugin-signed-with-sha256-certificate/
            // http://blogs.msdn.com/b/vsto/archive/2014/04/10/vsto-runtime-update-to-address-slow-shutdown-and-unknown-publisher-for-sha256-certificates.aspx

            outlookApp = Application;

            _inspectors = Application.Inspectors;
            _inspectors.NewInspector += _inspectors_NewInspector;
        }

        void _inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            if (Inspector == null) throw new ArgumentNullException("Inspector is null");

            theCurrentMailItem = null;

            object item = Inspector.CurrentItem;
            if (item == null) return;

            if (!(item is Outlook.MailItem)) return;

            theCurrentMailItem = Inspector.CurrentItem as Outlook.MailItem;

            ThisAddIn.bPeppermintMessageInserted = false;

        }

        public static DialogResult RecordAudioAndAttach(string RibbonName)
        {
            frmRecordAudio myRecordAudioForm = new frmRecordAudio();
            DialogResult dr = myRecordAudioForm.ShowDialog();

            if (dr == DialogResult.OK)
            {
                if (RibbonName == "Explorer") // Do not use the E-mail as a new one will be created by the caller
                    return dr;

                if (ThisAddIn.theCurrentMailItem == null)
                    return dr;

                if (!bPeppermintMessageInserted)
                {
                    if (ThisAddIn.theCurrentMailItem.Subject == null)
                    {
                        //if (RibbonName != "Read")
                        ThisAddIn.theCurrentMailItem.Subject = PEPPERMINT_NEW_MAIL_SUBJECT;
                    }
                    else
                    {
                        if ((String.IsNullOrEmpty(ThisAddIn.theCurrentMailItem.Subject.ToString())) && (RibbonName == "Create"))
                        {
                            ThisAddIn.theCurrentMailItem.Subject = PEPPERMINT_NEW_MAIL_SUBJECT;
                        }
                    }

                    if (ThisAddIn.theCurrentMailItem.Body != null)
                    {
                        // if the text "Peppermint.com" is in the body of the message don't update the body
                        if (!ThisAddIn.theCurrentMailItem.HTMLBody.ToString().Contains(PEPPERMINT_WEBSITE))
                        { 
                            ThisAddIn.theCurrentMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                            if ((!String.IsNullOrEmpty(ThisAddIn.theCurrentMailItem.Body.ToString())) && (RibbonName == "Create"))
                            {
                                ThisAddIn.theCurrentMailItem.HTMLBody = PEPPERMINT_NEW_EMAIL_HTML_BODY + ThisAddIn.theCurrentMailItem.HTMLBody;
                                bPeppermintMessageInserted = true;
                            }
                            if ((!String.IsNullOrEmpty(ThisAddIn.theCurrentMailItem.Body.ToString())) && (RibbonName == "Read"))
                            {
                                ThisAddIn.theCurrentMailItem.HTMLBody = PEPPERMINT_REPLY_MAIL_HTML_BODY + ThisAddIn.theCurrentMailItem.HTMLBody;
                                bPeppermintMessageInserted = true;
                            }
                        }
                    }
                }
                // Attach audio recording file
                if ((ThisAddIn.theCurrentMailItem != null) && (File.Exists(ThisAddIn.AttachmentFilePath)))
                    ThisAddIn.theCurrentMailItem.Attachments.Add(ThisAddIn.AttachmentFilePath);
            }

            return dr;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbons();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
