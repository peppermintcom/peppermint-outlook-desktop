using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Peppermint_Outlook_AddIn
{
    public partial class ReadEmailRibbon
    {
        private void ReadEmailRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnRecordMessage_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.MailItem mi = ThisAddIn.theCurrentMailItem.ReplyAll();
            mi.Display();
            ThisAddIn.RecordAudioAndAttach("Read");

        }
    }
}
