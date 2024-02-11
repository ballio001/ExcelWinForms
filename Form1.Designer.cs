
namespace ExcelWinForm
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
            this.CmdRead = new System.Windows.Forms.Button();
            this.CmdWrite = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.CmdPath = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CmdRead
            // 
            this.CmdRead.Location = new System.Drawing.Point(261, 96);
            this.CmdRead.Name = "CmdRead";
            this.CmdRead.Size = new System.Drawing.Size(85, 23);
            this.CmdRead.TabIndex = 0;
            this.CmdRead.Text = "Read";
            this.CmdRead.UseVisualStyleBackColor = true;
            this.CmdRead.Click += new System.EventHandler(this.CmdRead_Click);
            // 
            // CmdWrite
            // 
            this.CmdWrite.Location = new System.Drawing.Point(361, 95);
            this.CmdWrite.Name = "CmdWrite";
            this.CmdWrite.Size = new System.Drawing.Size(85, 24);
            this.CmdWrite.TabIndex = 1;
            this.CmdWrite.Text = "Write";
            this.CmdWrite.UseVisualStyleBackColor = true;
            this.CmdWrite.Click += new System.EventHandler(this.CmdWrite_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(13, 64);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(333, 23);
            this.textBox1.TabIndex = 2;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // CmdPath
            // 
            this.CmdPath.Location = new System.Drawing.Point(361, 64);
            this.CmdPath.Name = "CmdPath";
            this.CmdPath.Size = new System.Drawing.Size(85, 25);
            this.CmdPath.TabIndex = 3;
            this.CmdPath.Text = "Browse file";
            this.CmdPath.UseVisualStyleBackColor = true;
            this.CmdPath.Click += new System.EventHandler(this.CmdPath_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(477, 431);
            this.Controls.Add(this.CmdPath);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.CmdWrite);
            this.Controls.Add(this.CmdRead);
            this.Name = "Excel Manipulator";
            this.Text = "Excel Manipulator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button CmdRead;
        private System.Windows.Forms.Button CmdWrite;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button CmdPath;
    }
}