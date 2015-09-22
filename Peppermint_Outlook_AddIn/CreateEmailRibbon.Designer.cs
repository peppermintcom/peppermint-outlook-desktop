namespace Peppermint_Outlook_AddIn
{
    partial class CreateEmailRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CreateEmailRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnRecordMessage = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabNewMailMessage";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabNewMailMessage";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnRecordMessage);
            this.group1.Label = "Peppermint";
            this.group1.Name = "group1";
            this.group1.Position = this.Factory.RibbonPosition.BeforeOfficeId("GroupMessageOptions");
            // 
            // btnRecordMessage
            // 
            this.btnRecordMessage.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRecordMessage.Image = global::Peppermint_Outlook_AddIn.Properties.Resources.Logo;
            this.btnRecordMessage.Label = "Add Audio Message";
            this.btnRecordMessage.Name = "btnRecordMessage";
            this.btnRecordMessage.ShowImage = true;
            this.btnRecordMessage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRecordMessage_Click);
            // 
            // CreateEmailRibbon
            // 
            this.Name = "CreateEmailRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CreateEmailRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRecordMessage;
    }

    partial class ThisRibbonCollection
    {
        internal CreateEmailRibbon CreateEmailRibbon
        {
            get { return this.GetRibbon<CreateEmailRibbon>(); }
        }
    }
}
