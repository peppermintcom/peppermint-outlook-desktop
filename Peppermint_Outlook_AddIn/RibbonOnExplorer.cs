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

        private void btnSendViaPeppermint_Click(object sender, RibbonControlEventArgs e)
        {
            // Create a new email only of the audio is to be attached/sent
            if (ThisAddIn.RecordAudioAndAttach("Explorer") == DialogResult.OK )
            {
                ThisAddIn.theCurrentMailItem = ThisAddIn.outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                ThisAddIn.theCurrentMailItem.Display();

                ThisAddIn.theCurrentMailItem.Subject = "I sent you a voicemail message";
                
                ThisAddIn.theCurrentMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                ThisAddIn.theCurrentMailItem.HTMLBody = ThisAddIn.PEPPERMINT_REPLY_HTML_BODY + ThisAddIn.theCurrentMailItem.HTMLBody;
                ThisAddIn.bPeppermintMessageInserted = true;

                // Attach audio recording file
                if ((ThisAddIn.theCurrentMailItem != null) && (File.Exists(ThisAddIn.AttachmentFilePath)))
                    ThisAddIn.theCurrentMailItem.Attachments.Add(ThisAddIn.AttachmentFilePath);

            }
            
        }
    }
}
