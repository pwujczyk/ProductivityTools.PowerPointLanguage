namespace ProductivityTools.PowerPointLanguage
{
    partial class ChangeLanguage : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ChangeLanguage()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChangeLanguage));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.changeLanguageGroup = this.Factory.CreateRibbonGroup();
            this.slideButton = this.Factory.CreateRibbonButton();
            this.wholePresentation = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.changeLanguageGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.changeLanguageGroup);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // changeLanguageGroup
            // 
            this.changeLanguageGroup.Items.Add(this.slideButton);
            this.changeLanguageGroup.Items.Add(this.wholePresentation);
            this.changeLanguageGroup.Label = "Change Language";
            this.changeLanguageGroup.Name = "changeLanguageGroup";
            // 
            // slideButton
            // 
            this.slideButton.Image = ((System.Drawing.Image)(resources.GetObject("slideButton.Image")));
            this.slideButton.Label = "This slide";
            this.slideButton.Name = "slideButton";
            this.slideButton.ShowImage = true;
            // 
            // wholePresentation
            // 
            this.wholePresentation.Image = ((System.Drawing.Image)(resources.GetObject("wholePresentation.Image")));
            this.wholePresentation.Label = "All slides";
            this.wholePresentation.Name = "wholePresentation";
            this.wholePresentation.ShowImage = true;
            // 
            // ChangeLanguage
            // 
            this.Name = "ChangeLanguage";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ChangeLanguage_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.changeLanguageGroup.ResumeLayout(false);
            this.changeLanguageGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup changeLanguageGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton slideButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton wholePresentation;
    }

    partial class ThisRibbonCollection
    {
        internal ChangeLanguage ChangeLanguage
        {
            get { return this.GetRibbon<ChangeLanguage>(); }
        }
    }
}
