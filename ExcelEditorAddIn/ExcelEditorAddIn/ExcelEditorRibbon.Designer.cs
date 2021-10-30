
namespace ExcelEditorAddIn
{
    partial class ExcelEditorRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExcelEditorRibbon()
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
            this.JsonOpenButton = this.Factory.CreateRibbonButton();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.RecentsDropdown = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Excel Editor";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.JsonOpenButton);
            this.group1.Items.Add(this.RecentsDropdown);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // JsonOpenButton
            // 
            this.JsonOpenButton.Label = "Open";
            this.JsonOpenButton.Name = "JsonOpenButton";
            this.JsonOpenButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.JsonOpenButton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // RecentsDropdown
            // 
            this.RecentsDropdown.Label = "Recents";
            this.RecentsDropdown.Name = "RecentsDropdown";
            // 
            // ExcelEditorRibbon
            // 
            this.Name = "ExcelEditorRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExcelEditorRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton JsonOpenButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown RecentsDropdown;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelEditorRibbon ExcelEditorRibbon
        {
            get { return this.GetRibbon<ExcelEditorRibbon>(); }
        }
    }
}
