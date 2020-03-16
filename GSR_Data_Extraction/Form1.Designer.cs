namespace GSR_Data_Extraction
{
    partial class Form1
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
            this.openButton = new System.Windows.Forms.Button();
            this.convertButton = new System.Windows.Forms.Button();
            this.saveButton = new System.Windows.Forms.Button();
            this.idTextBox = new System.Windows.Forms.TextBox();
            this.startTimeTextBox = new System.Windows.Forms.TextBox();
            this.idLabel = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox = new System.Windows.Forms.GroupBox();
            this.statusLabel = new System.Windows.Forms.Label();
            this.statusRichTextBox = new System.Windows.Forms.RichTextBox();
            this.resetButton = new System.Windows.Forms.Button();
            this.groupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // openButton
            // 
            this.openButton.Location = new System.Drawing.Point(15, 12);
            this.openButton.Name = "openButton";
            this.openButton.Size = new System.Drawing.Size(134, 66);
            this.openButton.TabIndex = 0;
            this.openButton.Text = "Open";
            this.openButton.UseVisualStyleBackColor = true;
            this.openButton.Click += new System.EventHandler(this.openButton_Click);
            // 
            // convertButton
            // 
            this.convertButton.Location = new System.Drawing.Point(155, 12);
            this.convertButton.Name = "convertButton";
            this.convertButton.Size = new System.Drawing.Size(134, 66);
            this.convertButton.TabIndex = 1;
            this.convertButton.Text = "Convert";
            this.convertButton.UseVisualStyleBackColor = true;
            this.convertButton.Click += new System.EventHandler(this.convertButton_Click);
            // 
            // saveButton
            // 
            this.saveButton.Location = new System.Drawing.Point(295, 12);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(134, 66);
            this.saveButton.TabIndex = 2;
            this.saveButton.Text = "Save";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // idTextBox
            // 
            this.idTextBox.Location = new System.Drawing.Point(78, 99);
            this.idTextBox.Name = "idTextBox";
            this.idTextBox.Size = new System.Drawing.Size(100, 20);
            this.idTextBox.TabIndex = 3;
            this.idTextBox.TextChanged += new System.EventHandler(this.idTextBox_TextChanged);
            // 
            // startTimeTextBox
            // 
            this.startTimeTextBox.Location = new System.Drawing.Point(78, 125);
            this.startTimeTextBox.Name = "startTimeTextBox";
            this.startTimeTextBox.Size = new System.Drawing.Size(100, 20);
            this.startTimeTextBox.TabIndex = 4;
            this.startTimeTextBox.TextChanged += new System.EventHandler(this.startTimeTextBox_TextChanged);
            // 
            // idLabel
            // 
            this.idLabel.AutoSize = true;
            this.idLabel.Location = new System.Drawing.Point(9, 106);
            this.idLabel.Name = "idLabel";
            this.idLabel.Size = new System.Drawing.Size(25, 13);
            this.idLabel.TabIndex = 5;
            this.idLabel.Text = "ID#";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 132);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Start Time";
            // 
            // groupBox
            // 
            this.groupBox.Controls.Add(this.statusLabel);
            this.groupBox.Controls.Add(this.statusRichTextBox);
            this.groupBox.Location = new System.Drawing.Point(206, 84);
            this.groupBox.Name = "groupBox";
            this.groupBox.Size = new System.Drawing.Size(316, 73);
            this.groupBox.TabIndex = 7;
            this.groupBox.TabStop = false;
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(6, 15);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(37, 13);
            this.statusLabel.TabIndex = 9;
            this.statusLabel.Text = "Status";
            // 
            // statusRichTextBox
            // 
            this.statusRichTextBox.Location = new System.Drawing.Point(6, 31);
            this.statusRichTextBox.Name = "statusRichTextBox";
            this.statusRichTextBox.Size = new System.Drawing.Size(304, 30);
            this.statusRichTextBox.TabIndex = 8;
            this.statusRichTextBox.Text = "";
            // 
            // resetButton
            // 
            this.resetButton.Location = new System.Drawing.Point(435, 12);
            this.resetButton.Name = "resetButton";
            this.resetButton.Size = new System.Drawing.Size(81, 66);
            this.resetButton.TabIndex = 10;
            this.resetButton.Text = "Reset";
            this.resetButton.UseVisualStyleBackColor = true;
            this.resetButton.Click += new System.EventHandler(this.resetButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 163);
            this.Controls.Add(this.resetButton);
            this.Controls.Add(this.groupBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.idLabel);
            this.Controls.Add(this.startTimeTextBox);
            this.Controls.Add(this.idTextBox);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.convertButton);
            this.Controls.Add(this.openButton);
            this.Name = "Form1";
            this.Text = "GSR Data";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox.ResumeLayout(false);
            this.groupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button openButton;
        private System.Windows.Forms.Button convertButton;
        private System.Windows.Forms.Button saveButton;
        private System.Windows.Forms.TextBox idTextBox;
        private System.Windows.Forms.TextBox startTimeTextBox;
        private System.Windows.Forms.Label idLabel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.RichTextBox statusRichTextBox;
        private System.Windows.Forms.Button resetButton;
    }
}

