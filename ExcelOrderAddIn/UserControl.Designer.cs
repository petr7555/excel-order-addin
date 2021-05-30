
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
            this.createBtn = new System.Windows.Forms.Button();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
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
            this.idCol1ComboBox.FormattingEnabled = true;
            this.idCol1ComboBox.Location = new System.Drawing.Point(275, 47);
            this.idCol1ComboBox.Name = "idCol1ComboBox";
            this.idCol1ComboBox.Size = new System.Drawing.Size(121, 21);
            this.idCol1ComboBox.TabIndex = 3;
            this.idCol1ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.idCol1ComboBox_Validating);
            // 
            // idCol2ComboBox
            // 
            this.idCol2ComboBox.FormattingEnabled = true;
            this.idCol2ComboBox.Location = new System.Drawing.Point(275, 77);
            this.idCol2ComboBox.Name = "idCol2ComboBox";
            this.idCol2ComboBox.Size = new System.Drawing.Size(121, 21);
            this.idCol2ComboBox.TabIndex = 7;
            this.idCol2ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.idCol2ComboBox_Validating);
            // 
            // idCol2Label
            // 
            this.idCol2Label.AutoSize = true;
            this.idCol2Label.Location = new System.Drawing.Point(216, 80);
            this.idCol2Label.Name = "idCol2Label";
            this.idCol2Label.Size = new System.Drawing.Size(55, 13);
            this.idCol2Label.TabIndex = 6;
            this.idCol2Label.Text = "ID column";
            // 
            // table2Label
            // 
            this.table2Label.AutoSize = true;
            this.table2Label.Location = new System.Drawing.Point(11, 80);
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
            this.table2ComboBox.Location = new System.Drawing.Point(60, 77);
            this.table2ComboBox.Name = "table2ComboBox";
            this.table2ComboBox.Size = new System.Drawing.Size(121, 21);
            this.table2ComboBox.TabIndex = 4;
            this.table2ComboBox.SelectedIndexChanged += new System.EventHandler(this.table2ComboBox_SelectedIndexChanged);
            this.table2ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.table2ComboBox_Validating);
            // 
            // idCol3ComboBox
            // 
            this.idCol3ComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.idCol3ComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.idCol3ComboBox.FormattingEnabled = true;
            this.idCol3ComboBox.Location = new System.Drawing.Point(275, 104);
            this.idCol3ComboBox.Name = "idCol3ComboBox";
            this.idCol3ComboBox.Size = new System.Drawing.Size(121, 21);
            this.idCol3ComboBox.TabIndex = 11;
            this.idCol3ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.idCol3ComboBox_Validating);
            // 
            // idCol3Label
            // 
            this.idCol3Label.AutoSize = true;
            this.idCol3Label.Location = new System.Drawing.Point(216, 107);
            this.idCol3Label.Name = "idCol3Label";
            this.idCol3Label.Size = new System.Drawing.Size(55, 13);
            this.idCol3Label.TabIndex = 10;
            this.idCol3Label.Text = "ID column";
            // 
            // table3Label
            // 
            this.table3Label.AutoSize = true;
            this.table3Label.Location = new System.Drawing.Point(11, 107);
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
            this.table3ComboBox.Location = new System.Drawing.Point(60, 104);
            this.table3ComboBox.Name = "table3ComboBox";
            this.table3ComboBox.Size = new System.Drawing.Size(121, 21);
            this.table3ComboBox.TabIndex = 8;
            this.table3ComboBox.SelectedIndexChanged += new System.EventHandler(this.table3ComboBox_SelectedIndexChanged);
            this.table3ComboBox.Validating += new System.ComponentModel.CancelEventHandler(this.table3ComboBox_Validating);
            // 
            // createBtn
            // 
            this.createBtn.Location = new System.Drawing.Point(304, 131);
            this.createBtn.Name = "createBtn";
            this.createBtn.Size = new System.Drawing.Size(75, 23);
            this.createBtn.TabIndex = 12;
            this.createBtn.Text = "&Create";
            this.createBtn.UseVisualStyleBackColor = true;
            this.createBtn.Click += new System.EventHandler(this.createBtn_Click);
            // 
            // errorProvider
            // 
            this.errorProvider.ContainerControl = this;
            // 
            // UserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.createBtn);
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
            this.Name = "UserControl";
            this.Size = new System.Drawing.Size(436, 509);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox table1ComboBox;
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
        private System.Windows.Forms.Button createBtn;
        private System.Windows.Forms.ErrorProvider errorProvider;
    }
}
