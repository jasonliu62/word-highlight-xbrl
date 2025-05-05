namespace highlight
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonAnnotateFact = this.Factory.CreateRibbonButton();
            this.toggleAnnotateFact = this.Factory.CreateRibbonToggleButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnReplaceDisclosure = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Highlight and Modify";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonAnnotateFact);
            this.group1.Items.Add(this.toggleAnnotateFact);
            this.group1.Label = "Modify";
            this.group1.Name = "group1";
            // 
            // buttonAnnotateFact
            // 
            this.buttonAnnotateFact.Label = "Annotate Fact";
            this.buttonAnnotateFact.Name = "buttonAnnotateFact";
            this.buttonAnnotateFact.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAnnotateFact_Click);
            // 
            // toggleAnnotateFact
            // 
            this.toggleAnnotateFact.Image = ((System.Drawing.Image)(resources.GetObject("toggleAnnotateFact.Image")));
            this.toggleAnnotateFact.Label = "Modify Mode";
            this.toggleAnnotateFact.Name = "toggleAnnotateFact";
            this.toggleAnnotateFact.ShowImage = true;
            this.toggleAnnotateFact.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleAnnotateFact_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnReplaceDisclosure);
            this.group2.Label = "Insert";
            this.group2.Name = "group2";
            // 
            // btnReplaceDisclosure
            // 
            this.btnReplaceDisclosure.Image = ((System.Drawing.Image)(resources.GetObject("btnReplaceDisclosure.Image")));
            this.btnReplaceDisclosure.Label = "Replace Disclosure";
            this.btnReplaceDisclosure.Name = "btnReplaceDisclosure";
            this.btnReplaceDisclosure.ShowImage = true;
            this.btnReplaceDisclosure.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReplaceDisclosureButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAnnotateFact;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleAnnotateFact;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnReplaceDisclosure;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
