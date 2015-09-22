﻿namespace Peppermint_Outlook_AddIn
{
    partial class RibbonOnExplorer : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonOnExplorer()
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
            this.btnSendViaPeppermint = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSendViaPeppermint);
            this.group1.Label = "Peppermint";
            this.group1.Name = "group1";
            this.group1.Position = this.Factory.RibbonPosition.BeforeOfficeId("GroupMailDelete");
            // 
            // btnSendViaPeppermint
            // 
            this.btnSendViaPeppermint.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSendViaPeppermint.Image = global::Peppermint_Outlook_AddIn.Properties.Resources.Logo;
            this.btnSendViaPeppermint.Label = "Reply via Peppermint";
            this.btnSendViaPeppermint.Name = "btnSendViaPeppermint";
            this.btnSendViaPeppermint.ShowImage = true;
            this.btnSendViaPeppermint.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSendViaPeppermint_Click);
            // 
            // RibbonOnExplorer
            // 
            this.Name = "RibbonOnExplorer";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonOnExplorer_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSendViaPeppermint;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonOnExplorer RibbonOnExplorer
        {
            get { return this.GetRibbon<RibbonOnExplorer>(); }
        }
    }
}
