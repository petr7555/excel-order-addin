
namespace ExcelOrderAddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.openSidebarBtn = this.Factory.CreateRibbonButton();
            this.infoGroup = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.infoGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.infoGroup);
            this.tab1.Label = "Order Add-In";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.openSidebarBtn);
            this.group1.Label = "Controls";
            this.group1.Name = "group1";
            // 
            // openSidebarBtn
            // 
            this.openSidebarBtn.Image = global::ExcelOrderAddIn.Properties.Resources.open_outline;
            this.openSidebarBtn.Label = "Open sidebar";
            this.openSidebarBtn.Name = "openSidebarBtn";
            this.openSidebarBtn.ShowImage = true;
            this.openSidebarBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openSidebarBtn_Click);
            // 
            // infoGroup
            // 
            this.infoGroup.Items.Add(this.label1);
            this.infoGroup.Items.Add(this.label2);
            this.infoGroup.Label = "Info";
            this.infoGroup.Name = "infoGroup";
            // 
            // label1
            // 
            this.label1.Label = "Version: 0.1.3";
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.Label = "Created by Petr Janík";
            this.label2.Name = "label2";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.infoGroup.ResumeLayout(false);
            this.infoGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup infoGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton openSidebarBtn;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
