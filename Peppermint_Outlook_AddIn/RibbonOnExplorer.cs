using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;

namespace Peppermint_Outlook_AddIn
{
    public partial class RibbonOnExplorer
    {
        private void RibbonOnExplorer_Load(object sender, RibbonUIEventArgs e)
        {

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
                ThisAddIn.theCurrentMailItem.HTMLBody = ThisAddIn.PEPPERMINT_NEW_EMAIL_HTML_BODY + ThisAddIn.theCurrentMailItem.HTMLBody;
                ThisAddIn.bPeppermintMessageInserted = true;

                // Attach audio recording file
                if ((ThisAddIn.theCurrentMailItem != null) && (File.Exists(ThisAddIn.AttachmentFilePath)))
                    ThisAddIn.theCurrentMailItem.Attachments.Add(ThisAddIn.AttachmentFilePath);
            }
        }
        private void btnSendViaPeppermint_Click(object sender, RibbonControlEventArgs e)
        {
            try 
            {
                if (ThisAddIn.outlookApp.ActiveExplorer().Selection == null) ;
            }
            catch ( Exception ex )
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
                MessageBox.Show("Please select a single e-mail to respond to via Peppermint","More than 1 e-mail selected",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            if (ThisAddIn.outlookApp.ActiveExplorer().Selection.Count == 1 )
                if (ThisAddIn.outlookApp.ActiveExplorer().Selection[1] is Outlook.MailItem)
                {
                    ThisAddIn.theCurrentMailItem = ThisAddIn.outlookApp.ActiveExplorer().Selection[1] as Outlook.MailItem;

                    Outlook.MailItem mi = ThisAddIn.theCurrentMailItem.ReplyAll();
                    mi.Display();
                    ThisAddIn.RecordAudioAndAttach("Read");
                }
        }
    }
}
