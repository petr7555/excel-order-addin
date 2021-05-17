
namespace ExcelOrderAddIn
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.box2 = this.Factory.CreateRibbonBox();
            this.table1Combo = this.Factory.CreateRibbonComboBox();
            this.table1ColCombo = this.Factory.CreateRibbonComboBox();
            this.box3 = this.Factory.CreateRibbonBox();
            this.table2Combo = this.Factory.CreateRibbonComboBox();
            this.table2ColCombo = this.Factory.CreateRibbonComboBox();
            this.infoGroup = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box2.SuspendLayout();
            this.box3.SuspendLayout();
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
            this.group1.Items.Add(this.box2);
            this.group1.Items.Add(this.box3);
            this.group1.Label = "Create";
            this.group1.Name = "group1";
            this.group1.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group1_DialogLauncherClick);
            // 
            // box2
            // 
            this.box2.Items.Add(this.table1Combo);
            this.box2.Items.Add(this.table1ColCombo);
            this.box2.Name = "box2";
            // 
            // table1Combo
            // 
            this.table1Combo.Label = "Table 1";
            this.table1Combo.Name = "table1Combo";
            this.table1Combo.Text = null;
            this.table1Combo.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBox1_TextChanged);
            // 
            // table1ColCombo
            // 
            this.table1ColCombo.Label = "Id col";
            this.table1ColCombo.Name = "table1ColCombo";
            this.table1ColCombo.Text = null;
            this.table1ColCombo.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBox2_TextChanged);
            // 
            // box3
            // 
            this.box3.Items.Add(this.table2Combo);
            this.box3.Items.Add(this.table2ColCombo);
            this.box3.Name = "box3";
            // 
            // table2Combo
            // 
            this.table2Combo.Label = "Table 2";
            this.table2Combo.Name = "table2Combo";
            this.table2Combo.Text = null;
            // 
            // table2ColCombo
            // 
            this.table2ColCombo.Label = "Id col";
            this.table2ColCombo.Name = "table2ColCombo";
            this.table2ColCombo.Text = null;
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
            this.label1.Label = "Version: 0.0.1";
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.Label = "Created by Petr Janík";
            this.label2.Name = "label2";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.infoGroup.ResumeLayout(false);
            this.infoGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox table1Combo;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup infoGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox table1ColCombo;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox table2Combo;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox table2ColCombo;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
