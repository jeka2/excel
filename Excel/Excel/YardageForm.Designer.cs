namespace Excel
{
    partial class YardageForm
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
            this.styleLabel = new System.Windows.Forms.Label();
            this.yardageLabel = new System.Windows.Forms.Label();
            this.radioButton = new System.Windows.Forms.RadioButton();
            this.yardageText = new System.Windows.Forms.TextBox();
            this.saveButton = new System.Windows.Forms.Button();
            this.twomButton = new System.Windows.Forms.RadioButton();
            this.onemButton = new System.Windows.Forms.RadioButton();
            this.nmButton = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // styleLabel
            // 
            this.styleLabel.AutoSize = true;
            this.styleLabel.Location = new System.Drawing.Point(117, 101);
            this.styleLabel.Name = "styleLabel";
            this.styleLabel.Size = new System.Drawing.Size(38, 13);
            this.styleLabel.TabIndex = 0;
            this.styleLabel.Text = "label1 ";
            // 
            // yardageLabel
            // 
            this.yardageLabel.AutoSize = true;
            this.yardageLabel.Location = new System.Drawing.Point(88, 175);
            this.yardageLabel.Name = "yardageLabel";
            this.yardageLabel.Size = new System.Drawing.Size(45, 13);
            this.yardageLabel.TabIndex = 1;
            this.yardageLabel.Text = "yardage";
            // 
            // radioButton
            // 
            this.radioButton.AutoSize = true;
            this.radioButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.radioButton.Location = new System.Drawing.Point(290, 204);
            this.radioButton.Name = "radioButton";
            this.radioButton.Size = new System.Drawing.Size(24, 5);
            this.radioButton.TabIndex = 2;
            this.radioButton.TabStop = true;
            this.radioButton.UseVisualStyleBackColor = true;
            // 
            // yardageText
            // 
            this.yardageText.Location = new System.Drawing.Point(153, 175);
            this.yardageText.Name = "yardageText";
            this.yardageText.Size = new System.Drawing.Size(99, 20);
            this.yardageText.TabIndex = 5;
            // 
            // saveButton
            // 
            this.saveButton.Location = new System.Drawing.Point(191, 235);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(107, 26);
            this.saveButton.TabIndex = 6;
            this.saveButton.Text = "Save";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // twomButton
            // 
            this.twomButton.AutoSize = true;
            this.twomButton.Location = new System.Drawing.Point(394, 178);
            this.twomButton.Name = "twomButton";
            this.twomButton.Size = new System.Drawing.Size(40, 17);
            this.twomButton.TabIndex = 7;
            this.twomButton.Text = "2M";
            this.twomButton.UseVisualStyleBackColor = true;
            // 
            // onemButton
            // 
            this.onemButton.AutoSize = true;
            this.onemButton.Location = new System.Drawing.Point(331, 178);
            this.onemButton.Name = "onemButton";
            this.onemButton.Size = new System.Drawing.Size(40, 17);
            this.onemButton.TabIndex = 3;
            this.onemButton.Text = "1M";
            this.onemButton.UseVisualStyleBackColor = true;
            // 
            // nmButton
            // 
            this.nmButton.AutoSize = true;
            this.nmButton.Location = new System.Drawing.Point(272, 178);
            this.nmButton.Name = "nmButton";
            this.nmButton.Size = new System.Drawing.Size(42, 17);
            this.nmButton.TabIndex = 4;
            this.nmButton.Text = "NM";
            this.nmButton.UseVisualStyleBackColor = true;
            // 
            // YardageForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(504, 287);
            this.Controls.Add(this.twomButton);
            this.Controls.Add(this.onemButton);
            this.Controls.Add(this.nmButton);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.yardageText);
            this.Controls.Add(this.radioButton);
            this.Controls.Add(this.yardageLabel);
            this.Controls.Add(this.styleLabel);
            this.Name = "YardageForm";
            this.Text = "YardageForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label styleLabel;
        private System.Windows.Forms.Label yardageLabel;
        private System.Windows.Forms.RadioButton radioButton;
        private System.Windows.Forms.TextBox yardageText;
        private System.Windows.Forms.Button saveButton;
        private System.Windows.Forms.RadioButton twomButton;
        private System.Windows.Forms.RadioButton onemButton;
        private System.Windows.Forms.RadioButton nmButton;
    }
}