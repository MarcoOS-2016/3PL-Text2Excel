namespace _3PL_Text2Excel
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
            this.outputFolderButton = new System.Windows.Forms.Button();
            this.sourceFolderButton = new System.Windows.Forms.Button();
            this.outputFolderTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.startButton = new System.Windows.Forms.Button();
            this.sourceFolderTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // outputFolderButton
            // 
            this.outputFolderButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.outputFolderButton.Image = ((System.Drawing.Image)(resources.GetObject("outputFolderButton.Image")));
            this.outputFolderButton.Location = new System.Drawing.Point(574, 64);
            this.outputFolderButton.Name = "outputFolderButton";
            this.outputFolderButton.Size = new System.Drawing.Size(28, 25);
            this.outputFolderButton.TabIndex = 50;
            this.outputFolderButton.UseVisualStyleBackColor = true;
            this.outputFolderButton.Click += new System.EventHandler(this.outputFolderButton_Click);
            // 
            // sourceFolderButton
            // 
            this.sourceFolderButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sourceFolderButton.Image = ((System.Drawing.Image)(resources.GetObject("sourceFolderButton.Image")));
            this.sourceFolderButton.Location = new System.Drawing.Point(574, 22);
            this.sourceFolderButton.Name = "sourceFolderButton";
            this.sourceFolderButton.Size = new System.Drawing.Size(28, 25);
            this.sourceFolderButton.TabIndex = 49;
            this.sourceFolderButton.UseVisualStyleBackColor = true;
            this.sourceFolderButton.Click += new System.EventHandler(this.sourceFolderButton_Click);
            // 
            // outputFolderTextBox
            // 
            this.outputFolderTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.outputFolderTextBox.Location = new System.Drawing.Point(179, 65);
            this.outputFolderTextBox.Name = "outputFolderTextBox";
            this.outputFolderTextBox.Size = new System.Drawing.Size(385, 23);
            this.outputFolderTextBox.TabIndex = 48;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(81, 66);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 16);
            this.label2.TabIndex = 47;
            this.label2.Text = "Output Folder:";
            // 
            // startButton
            // 
            this.startButton.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.startButton.Location = new System.Drawing.Point(503, 117);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(100, 30);
            this.startButton.TabIndex = 46;
            this.startButton.Text = "Start";
            this.startButton.UseVisualStyleBackColor = true;
            this.startButton.Click += new System.EventHandler(this.startButton_Click);
            // 
            // sourceFolderTextBox
            // 
            this.sourceFolderTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.sourceFolderTextBox.Location = new System.Drawing.Point(179, 23);
            this.sourceFolderTextBox.Name = "sourceFolderTextBox";
            this.sourceFolderTextBox.Size = new System.Drawing.Size(385, 23);
            this.sourceFolderTextBox.TabIndex = 45;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(16, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 16);
            this.label1.TabIndex = 44;
            this.label1.Text = "Folder of Source Reports:";
            // 
            // statusStrip
            // 
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel});
            this.statusStrip.Location = new System.Drawing.Point(0, 170);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(622, 22);
            this.statusStrip.TabIndex = 51;
            this.statusStrip.Text = "statusStrip1";
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(0, 17);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(622, 192);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.outputFolderButton);
            this.Controls.Add(this.sourceFolderButton);
            this.Controls.Add(this.outputFolderTextBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.startButton);
            this.Controls.Add(this.sourceFolderTextBox);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "3PL Text2Excel (Ver.07272015)";
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button outputFolderButton;
        private System.Windows.Forms.Button sourceFolderButton;
        private System.Windows.Forms.TextBox outputFolderTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button startButton;
        private System.Windows.Forms.TextBox sourceFolderTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
    }
}

