namespace HotREF
{
    partial class MainForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.buttonExcel = new System.Windows.Forms.Button();
            this.selectProposedFile = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.createRefButton = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ZoneSelectBox = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.TemplateButton = new System.Windows.Forms.Button();
            this.CreatePropButton = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.TemplateDirectoryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.worksheetTextBox = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // buttonExcel
            // 
            this.buttonExcel.Location = new System.Drawing.Point(12, 36);
            this.buttonExcel.Name = "buttonExcel";
            this.buttonExcel.Size = new System.Drawing.Size(335, 23);
            this.buttonExcel.TabIndex = 0;
            this.buttonExcel.Text = "Choose Worksheet";
            this.buttonExcel.UseVisualStyleBackColor = true;
            this.buttonExcel.Click += new System.EventHandler(this.Choose_Worksheet_Button);
            // 
            // selectProposedFile
            // 
            this.selectProposedFile.Location = new System.Drawing.Point(19, 19);
            this.selectProposedFile.Name = "selectProposedFile";
            this.selectProposedFile.Size = new System.Drawing.Size(209, 23);
            this.selectProposedFile.TabIndex = 3;
            this.selectProposedFile.Text = "Select Proposed File";
            this.selectProposedFile.UseVisualStyleBackColor = true;
            this.selectProposedFile.Click += new System.EventHandler(this.SelectProposedFile_Click);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(18, 81);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(228, 13);
            this.textBox1.TabIndex = 4;
            this.textBox1.Text = "No file selected";
            // 
            // createRefButton
            // 
            this.createRefButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.createRefButton.Location = new System.Drawing.Point(19, 118);
            this.createRefButton.Name = "createRefButton";
            this.createRefButton.Size = new System.Drawing.Size(311, 23);
            this.createRefButton.TabIndex = 12;
            this.createRefButton.Text = "Create REF";
            this.createRefButton.UseVisualStyleBackColor = true;
            this.createRefButton.Click += new System.EventHandler(this.Button1_Click);
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(15, 59);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 19);
            this.label4.TabIndex = 1;
            this.label4.Text = "Proposed File:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ZoneSelectBox);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.createRefButton);
            this.groupBox1.Controls.Add(this.selectProposedFile);
            this.groupBox1.Location = new System.Drawing.Point(12, 202);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(336, 147);
            this.groupBox1.TabIndex = 20;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "REFERENCE";
            // 
            // ZoneSelectBox
            // 
            this.ZoneSelectBox.FormattingEnabled = true;
            this.ZoneSelectBox.Items.AddRange(new object[] {
            "Zone 7A",
            "Zone 6"});
            this.ZoneSelectBox.Location = new System.Drawing.Point(241, 19);
            this.ZoneSelectBox.Name = "ZoneSelectBox";
            this.ZoneSelectBox.Size = new System.Drawing.Size(89, 21);
            this.ZoneSelectBox.TabIndex = 0;
            this.ZoneSelectBox.Text = "Zone 7A";
            this.ZoneSelectBox.SelectedIndexChanged += new System.EventHandler(this.ZoneSelectBox_SelectedIndexChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.TemplateButton);
            this.groupBox2.Controls.Add(this.CreatePropButton);
            this.groupBox2.Location = new System.Drawing.Point(12, 106);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(335, 90);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "PROPOSED";
            // 
            // TemplateButton
            // 
            this.TemplateButton.Location = new System.Drawing.Point(6, 19);
            this.TemplateButton.Name = "TemplateButton";
            this.TemplateButton.Size = new System.Drawing.Size(324, 23);
            this.TemplateButton.TabIndex = 1;
            this.TemplateButton.Text = "Choose Template";
            this.TemplateButton.UseVisualStyleBackColor = true;
            this.TemplateButton.Click += new System.EventHandler(this.TemplateButton_Click);
            // 
            // CreatePropButton
            // 
            this.CreatePropButton.Location = new System.Drawing.Point(6, 48);
            this.CreatePropButton.Name = "CreatePropButton";
            this.CreatePropButton.Size = new System.Drawing.Size(324, 23);
            this.CreatePropButton.TabIndex = 2;
            this.CreatePropButton.Text = "Create House From Template";
            this.CreatePropButton.UseVisualStyleBackColor = true;
            this.CreatePropButton.Click += new System.EventHandler(this.CreateProp_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.settingsToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(359, 24);
            this.menuStrip1.TabIndex = 21;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.TemplateDirectoryToolStripMenuItem});
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.settingsToolStripMenuItem.Text = "Settings";
            // 
            // TemplateDirectoryToolStripMenuItem
            // 
            this.TemplateDirectoryToolStripMenuItem.Name = "TemplateDirectoryToolStripMenuItem";
            this.TemplateDirectoryToolStripMenuItem.Size = new System.Drawing.Size(212, 22);
            this.TemplateDirectoryToolStripMenuItem.Text = "Default template directory";
            this.TemplateDirectoryToolStripMenuItem.Click += new System.EventHandler(this.TemplateDirectoryToolStripMenuItem_Click);
            // 
            // worksheetTextBox
            // 
            this.worksheetTextBox.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.worksheetTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.worksheetTextBox.Location = new System.Drawing.Point(12, 65);
            this.worksheetTextBox.Name = "worksheetTextBox";
            this.worksheetTextBox.ReadOnly = true;
            this.worksheetTextBox.Size = new System.Drawing.Size(228, 13);
            this.worksheetTextBox.TabIndex = 22;
            this.worksheetTextBox.Text = "No file selected";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(359, 361);
            this.Controls.Add(this.worksheetTextBox);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonExcel);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimumSize = new System.Drawing.Size(365, 300);
            this.Name = "MainForm";
            this.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Text = "HotREF";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button buttonExcel;
        private System.Windows.Forms.Button selectProposedFile;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button createRefButton;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button CreatePropButton;
        private System.Windows.Forms.Button TemplateButton;
        private System.Windows.Forms.ComboBox ZoneSelectBox;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem TemplateDirectoryToolStripMenuItem;
        private System.Windows.Forms.TextBox worksheetTextBox;
    }
}

