using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Peppermint_Outlook_AddIn
{
    public partial class CreateEmailRibbon
    {
        private void CreateEmailRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnRecordMessage_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.theCurrentMailItem.Save();
            ThisAddIn.RecordAudioAndAttach("Create");
        }
    }
}
