﻿
namespace ExcelOrderAddIn
{
    partial class UserControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UserControl));
            this.table1ComboBox = new System.Windows.Forms.ComboBox();
            this.table1Label = new System.Windows.Forms.Label();
            this.idCol1Label = new System.Windows.Forms.Label();
            this.idCol1ComboBox = new System.Windows.Forms.ComboBox();
            this.idCol2ComboBox = new System.Windows.Forms.ComboBox();
            this.idCol2Label = new System.Windows.Forms.Label();
            this.table2Label = new System.Windows.Forms.Label();
            this.table2ComboBox = new System.Windows.Forms.ComboBox();
            this.idCol3ComboBox = new System.Windows.Forms.ComboBox();
            this.idCol3Label = new System.Windows.Forms.Label();
            this.table3Label = new System.Windows.Forms.Label();
            this.table3ComboBox = new System.Windows.Forms.ComboBox();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.deleteNotGeneratedSheetsBtn = new System.Windows.Forms.Button();
            this.deleteGeneratedSheetsBtn = new System.Windows.Forms.Button();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.selectImgFolderBtn = new System.Windows.Forms.Button();
            this.createBtn = new System.Windows.Forms.Button();
            this.refreshBtn = new System.Windows.Forms.Button();
            this.refreshBtnImageList = new System.Windows.Forms.ImageList(this.components);
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.imgFolderTextBox = new System.Windows.Forms.TextBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.progressBarLabel = new System.Windows.Forms.Label();
            this.logsLabel = new System.Windows.Forms.Label();
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.scrollingRichTextBox = new ExcelOrderAddIn.ScrollingRichTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // table1ComboBox
            // 
            this.table1ComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.table1ComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.table1ComboBox.FormattingEnabled = true;
            this.table1ComboBox.Location = new System.Drawing.Point(60, 47);
            this.table1ComboBox.Name = "table1ComboBox";
            this.table1ComboBox.Size = new System.Drawing.Size(121, 21);
            this.table1ComboBox.TabIndex = 0;
            this.toolTip.SetToolTip(this.table1ComboBox, "Select first table.");
            this.table1ComboBox.SelectedIndexChanged += new System.EventHandler(this.table1ComboBox_SelectedIndexChanged);
            this.table1ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.table1ComboBox_Validating);
            // 
            // table1Label
            // 
            this.table1Label.AutoSize = true;
            this.table1Label.Location = new System.Drawing.Point(11, 50);
            this.table1Label.Name = "table1Label";
            this.table1Label.Size = new System.Drawing.Size(43, 13);
            this.table1Label.TabIndex = 1;
            this.table1Label.Text = "Table 1";
            // 
            // idCol1Label
            // 
            this.idCol1Label.AutoSize = true;
            this.idCol1Label.Location = new System.Drawing.Point(216, 50);
            this.idCol1Label.Name = "idCol1Label";
            this.idCol1Label.Size = new System.Drawing.Size(55, 13);
            this.idCol1Label.TabIndex = 2;
            this.idCol1Label.Text = "ID column";
            // 
            // idCol1ComboBox
            // 
            this.idCol1ComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.idCol1ComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.idCol1ComboBox.FormattingEnabled = true;
            this.idCol1ComboBox.Location = new System.Drawing.Point(275, 47);
            this.idCol1ComboBox.Name = "idCol1ComboBox";
            this.idCol1ComboBox.Size = new System.Drawing.Size(121, 21);
            this.idCol1ComboBox.TabIndex = 3;
            this.toolTip.SetToolTip(this.idCol1ComboBox, "Common column to join on.");
            this.idCol1ComboBox.SelectedIndexChanged += new System.EventHandler(this.idCol1ComboBox_SelectedIndexChanged);
            this.idCol1ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.idCol1ComboBox_Validating);
            // 
            // idCol2ComboBox
            // 
            this.idCol2ComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.idCol2ComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.idCol2ComboBox.FormattingEnabled = true;
            this.idCol2ComboBox.Location = new System.Drawing.Point(275, 74);
            this.idCol2ComboBox.Name = "idCol2ComboBox";
            this.idCol2ComboBox.Size = new System.Drawing.Size(121, 21);
            this.idCol2ComboBox.TabIndex = 7;
            this.toolTip.SetToolTip(this.idCol2ComboBox, "Common column to join on.");
            this.idCol2ComboBox.SelectedIndexChanged += new System.EventHandler(this.idCol2ComboBox_SelectedIndexChanged);
            this.idCol2ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.idCol2ComboBox_Validating);
            // 
            // idCol2Label
            // 
            this.idCol2Label.AutoSize = true;
            this.idCol2Label.Location = new System.Drawing.Point(216, 77);
            this.idCol2Label.Name = "idCol2Label";
            this.idCol2Label.Size = new System.Drawing.Size(55, 13);
            this.idCol2Label.TabIndex = 6;
            this.idCol2Label.Text = "ID column";
            // 
            // table2Label
            // 
            this.table2Label.AutoSize = true;
            this.table2Label.Location = new System.Drawing.Point(11, 77);
            this.table2Label.Name = "table2Label";
            this.table2Label.Size = new System.Drawing.Size(43, 13);
            this.table2Label.TabIndex = 5;
            this.table2Label.Text = "Table 2";
            // 
            // table2ComboBox
            // 
            this.table2ComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.table2ComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.table2ComboBox.FormattingEnabled = true;
            this.table2ComboBox.Location = new System.Drawing.Point(60, 74);
            this.table2ComboBox.Name = "table2ComboBox";
            this.table2ComboBox.Size = new System.Drawing.Size(121, 21);
            this.table2ComboBox.TabIndex = 4;
            this.toolTip.SetToolTip(this.table2ComboBox, "Select second table.");
            this.table2ComboBox.SelectedIndexChanged += new System.EventHandler(this.table2ComboBox_SelectedIndexChanged);
            this.table2ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.table2ComboBox_Validating);
            // 
            // idCol3ComboBox
            // 
            this.idCol3ComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.idCol3ComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.idCol3ComboBox.FormattingEnabled = true;
            this.idCol3ComboBox.Location = new System.Drawing.Point(275, 101);
            this.idCol3ComboBox.Name = "idCol3ComboBox";
            this.idCol3ComboBox.Size = new System.Drawing.Size(121, 21);
            this.idCol3ComboBox.TabIndex = 11;
            this.toolTip.SetToolTip(this.idCol3ComboBox, "Common column to join on.");
            this.idCol3ComboBox.SelectedIndexChanged += new System.EventHandler(this.idCol3ComboBox_SelectedIndexChanged);
            this.idCol3ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.idCol3ComboBox_Validating);
            // 
            // idCol3Label
            // 
            this.idCol3Label.AutoSize = true;
            this.idCol3Label.Location = new System.Drawing.Point(216, 104);
            this.idCol3Label.Name = "idCol3Label";
            this.idCol3Label.Size = new System.Drawing.Size(55, 13);
            this.idCol3Label.TabIndex = 10;
            this.idCol3Label.Text = "ID column";
            // 
            // table3Label
            // 
            this.table3Label.AutoSize = true;
            this.table3Label.Location = new System.Drawing.Point(11, 104);
            this.table3Label.Name = "table3Label";
            this.table3Label.Size = new System.Drawing.Size(43, 13);
            this.table3Label.TabIndex = 9;
            this.table3Label.Text = "Table 3";
            // 
            // table3ComboBox
            // 
            this.table3ComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.table3ComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.table3ComboBox.FormattingEnabled = true;
            this.table3ComboBox.Location = new System.Drawing.Point(60, 101);
            this.table3ComboBox.Name = "table3ComboBox";
            this.table3ComboBox.Size = new System.Drawing.Size(121, 21);
            this.table3ComboBox.TabIndex = 8;
            this.toolTip.SetToolTip(this.table3ComboBox, "Select third table.");
            this.table3ComboBox.SelectedIndexChanged += new System.EventHandler(this.table3ComboBox_SelectedIndexChanged);
            this.table3ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.table3ComboBox_Validating);
            // 
            // errorProvider
            // 
            this.errorProvider.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
            this.errorProvider.ContainerControl = this;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox1.Controls.Add(this.deleteNotGeneratedSheetsBtn);
            this.groupBox1.Controls.Add(this.deleteGeneratedSheetsBtn);
            this.groupBox1.Location = new System.Drawing.Point(14, 367);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 100);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Danger zone";
            // 
            // deleteNotGeneratedSheetsBtn
            // 
            this.deleteNotGeneratedSheetsBtn.BackColor = System.Drawing.Color.LightCoral;
            this.deleteNotGeneratedSheetsBtn.Location = new System.Drawing.Point(6, 48);
            this.deleteNotGeneratedSheetsBtn.Name = "deleteNotGeneratedSheetsBtn";
            this.deleteNotGeneratedSheetsBtn.Size = new System.Drawing.Size(188, 23);
            this.deleteNotGeneratedSheetsBtn.TabIndex = 1;
            this.deleteNotGeneratedSheetsBtn.Text = "Delete not generated sheets";
            this.toolTip.SetToolTip(this.deleteNotGeneratedSheetsBtn, "Deletes all sheets NOT starting with \'New Order\'.");
            this.deleteNotGeneratedSheetsBtn.UseVisualStyleBackColor = false;
            this.deleteNotGeneratedSheetsBtn.Click += new System.EventHandler(this.deleteNotGeneratedSheetsBtn_Click);
            // 
            // deleteGeneratedSheetsBtn
            // 
            this.deleteGeneratedSheetsBtn.BackColor = System.Drawing.Color.LightCoral;
            this.deleteGeneratedSheetsBtn.Location = new System.Drawing.Point(6, 19);
            this.deleteGeneratedSheetsBtn.Name = "deleteGeneratedSheetsBtn";
            this.deleteGeneratedSheetsBtn.Size = new System.Drawing.Size(188, 23);
            this.deleteGeneratedSheetsBtn.TabIndex = 0;
            this.deleteGeneratedSheetsBtn.Text = "Delete generated sheets";
            this.toolTip.SetToolTip(this.deleteGeneratedSheetsBtn, "Deletes all sheets starting with \'New Order\'.");
            this.deleteGeneratedSheetsBtn.UseVisualStyleBackColor = false;
            this.deleteGeneratedSheetsBtn.Click += new System.EventHandler(this.deleteGeneratedSheetsBtn_Click);
            // 
            // selectImgFolderBtn
            // 
            this.selectImgFolderBtn.Location = new System.Drawing.Point(275, 128);
            this.selectImgFolderBtn.Name = "selectImgFolderBtn";
            this.selectImgFolderBtn.Size = new System.Drawing.Size(121, 23);
            this.selectImgFolderBtn.TabIndex = 14;
            this.selectImgFolderBtn.Text = "&Select images folder";
            this.toolTip.SetToolTip(this.selectImgFolderBtn, "Folder to look for the images in.");
            this.selectImgFolderBtn.UseVisualStyleBackColor = true;
            this.selectImgFolderBtn.Click += new System.EventHandler(this.selectImgFolderBtn_Click);
            // 
            // createBtn
            // 
            this.createBtn.BackColor = System.Drawing.Color.LawnGreen;
            this.createBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.createBtn.Location = new System.Drawing.Point(314, 157);
            this.createBtn.Name = "createBtn";
            this.createBtn.Size = new System.Drawing.Size(82, 29);
            this.createBtn.TabIndex = 20;
            this.createBtn.Text = "&Create";
            this.toolTip.SetToolTip(this.createBtn, "Creates  \'New Order\' sheet.");
            this.createBtn.UseVisualStyleBackColor = false;
            this.createBtn.Click += new System.EventHandler(this.createBtn_Click);
            // 
            // refreshBtn
            // 
            this.refreshBtn.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.refreshBtn.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.refreshBtn.ImageIndex = 0;
            this.refreshBtn.ImageList = this.refreshBtnImageList;
            this.refreshBtn.Location = new System.Drawing.Point(308, 12);
            this.refreshBtn.Name = "refreshBtn";
            this.refreshBtn.Size = new System.Drawing.Size(88, 29);
            this.refreshBtn.TabIndex = 17;
            this.refreshBtn.Text = "&Refresh";
            this.refreshBtn.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.toolTip.SetToolTip(this.refreshBtn, "Updates available tables in combo boxes. Click this when you create or delete a s" +
        "heet.");
            this.refreshBtn.UseVisualStyleBackColor = false;
            this.refreshBtn.Click += new System.EventHandler(this.refreshBtn_Click);
            // 
            // refreshBtnImageList
            // 
            this.refreshBtnImageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("refreshBtnImageList.ImageStream")));
            this.refreshBtnImageList.TransparentColor = System.Drawing.Color.Transparent;
            this.refreshBtnImageList.Images.SetKeyName(0, "refresh.png");
            // 
            // imgFolderTextBox
            // 
            this.imgFolderTextBox.Location = new System.Drawing.Point(14, 130);
            this.imgFolderTextBox.Name = "imgFolderTextBox";
            this.imgFolderTextBox.ReadOnly = true;
            this.imgFolderTextBox.Size = new System.Drawing.Size(257, 20);
            this.imgFolderTextBox.TabIndex = 15;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(14, 157);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(294, 29);
            this.progressBar.TabIndex = 19;
            // 
            // progressBarLabel
            // 
            this.progressBarLabel.AutoSize = true;
            this.progressBarLabel.Location = new System.Drawing.Point(11, 190);
            this.progressBarLabel.Name = "progressBarLabel";
            this.progressBarLabel.Size = new System.Drawing.Size(0, 13);
            this.progressBarLabel.TabIndex = 21;
            // 
            // logsLabel
            // 
            this.logsLabel.AutoSize = true;
            this.logsLabel.Location = new System.Drawing.Point(11, 217);
            this.logsLabel.Name = "logsLabel";
            this.logsLabel.Size = new System.Drawing.Size(33, 13);
            this.logsLabel.TabIndex = 23;
            this.logsLabel.Text = "Logs:";
            // 
            // timer
            // 
            this.timer.Interval = 500;
            this.timer.Tick += new System.EventHandler(this.TimeUpdateLogWindow_Tick);
            // 
            // scrollingRichTextBox
            // 
            this.scrollingRichTextBox.Location = new System.Drawing.Point(14, 234);
            this.scrollingRichTextBox.Name = "scrollingRichTextBox";
            this.scrollingRichTextBox.ReadOnly = true;
            this.scrollingRichTextBox.Size = new System.Drawing.Size(382, 127);
            this.scrollingRichTextBox.TabIndex = 24;
            this.scrollingRichTextBox.Text = "";
            // 
            // UserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.Disable;
            this.Controls.Add(this.scrollingRichTextBox);
            this.Controls.Add(this.logsLabel);
            this.Controls.Add(this.progressBarLabel);
            this.Controls.Add(this.createBtn);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.refreshBtn);
            this.Controls.Add(this.imgFolderTextBox);
            this.Controls.Add(this.selectImgFolderBtn);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.idCol3ComboBox);
            this.Controls.Add(this.idCol3Label);
            this.Controls.Add(this.table3Label);
            this.Controls.Add(this.table3ComboBox);
            this.Controls.Add(this.idCol2ComboBox);
            this.Controls.Add(this.idCol2Label);
            this.Controls.Add(this.table2Label);
            this.Controls.Add(this.table2ComboBox);
            this.Controls.Add(this.idCol1ComboBox);
            this.Controls.Add(this.idCol1Label);
            this.Controls.Add(this.table1Label);
            this.Controls.Add(this.table1ComboBox);
            this.MinimumSize = new System.Drawing.Size(408, 300);
            this.Name = "UserControl";
            this.Size = new System.Drawing.Size(413, 470);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label table1Label;
        private System.Windows.Forms.Label idCol1Label;
        private System.Windows.Forms.ComboBox idCol1ComboBox;
        private System.Windows.Forms.ComboBox idCol2ComboBox;
        private System.Windows.Forms.Label idCol2Label;
        private System.Windows.Forms.Label table2Label;
        private System.Windows.Forms.ComboBox table2ComboBox;
        private System.Windows.Forms.ComboBox idCol3ComboBox;
        private System.Windows.Forms.Label idCol3Label;
        private System.Windows.Forms.Label table3Label;
        private System.Windows.Forms.ComboBox table3ComboBox;
        private System.Windows.Forms.ErrorProvider errorProvider;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button deleteGeneratedSheetsBtn;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.TextBox imgFolderTextBox;
        private System.Windows.Forms.Button selectImgFolderBtn;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Button refreshBtn;
        private System.Windows.Forms.ImageList refreshBtnImageList;
        private System.Windows.Forms.ComboBox table1ComboBox;
        private System.Windows.Forms.Button createBtn;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label progressBarLabel;
        private System.Windows.Forms.Button deleteNotGeneratedSheetsBtn;
        private System.Windows.Forms.Label logsLabel;
        private ScrollingRichTextBox scrollingRichTextBox;
        private System.Windows.Forms.Timer timer;
    }
}
