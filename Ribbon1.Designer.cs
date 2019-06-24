namespace LCARenamerAddIn
{
    partial class TestRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TestRibbon()
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
            this.TestGroup = this.Factory.CreateRibbonGroup();
            this.TestButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.TestGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.TestGroup);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // TestGroup
            // 
            this.TestGroup.Items.Add(this.TestButton);
            this.TestGroup.Label = "TestGroup";
            this.TestGroup.Name = "TestGroup";
            // 
            // TestButton
            // 
            this.TestButton.Label = "TestButton";
            this.TestButton.Name = "TestButton";
            this.TestButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestButton_Click);
            // 
            // TestRibbon
            // 
            this.Name = "TestRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.TestGroup.ResumeLayout(false);
            this.TestGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup TestGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestButton;
    }

    partial class ThisRibbonCollection
    {
        internal TestRibbon Ribbon1
        {
            get { return this.GetRibbon<TestRibbon>(); }
        }
    }
}
