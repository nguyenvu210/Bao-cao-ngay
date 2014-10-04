using Process = System.Diagnostics.Process;
namespace Tool_Report
{
    partial class ToolReport
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
            Process[] prs = Process.GetProcesses();
            foreach (Process pr in prs)
            {
                if (pr.ProcessName == "EXCEL")
                {
                    pr.Kill();
                }
            }
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ToolReport));
            this.button1 = new System.Windows.Forms.Button();
            this.Template = new System.Windows.Forms.Button();
            this.Run = new System.Windows.Forms.Button();
            this.DR2G = new System.Windows.Forms.RadioButton();
            this.DR3G = new System.Windows.Forms.RadioButton();
            this.lst_logfile = new System.Windows.Forms.ListBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(13, 10);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Data";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Data_Click);
            // 
            // Template
            // 
            this.Template.Location = new System.Drawing.Point(13, 39);
            this.Template.Name = "Template";
            this.Template.Size = new System.Drawing.Size(75, 23);
            this.Template.TabIndex = 1;
            this.Template.Text = "Template";
            this.Template.UseVisualStyleBackColor = true;
            this.Template.Click += new System.EventHandler(this.Template_Click);
            // 
            // Run
            // 
            this.Run.Location = new System.Drawing.Point(13, 236);
            this.Run.Name = "Run";
            this.Run.Size = new System.Drawing.Size(75, 23);
            this.Run.TabIndex = 2;
            this.Run.Text = "Run";
            this.Run.UseVisualStyleBackColor = true;
            this.Run.Click += new System.EventHandler(this.Run_Click);
            // 
            // DR2G
            // 
            this.DR2G.AutoSize = true;
            this.DR2G.BackColor = System.Drawing.Color.Transparent;
            this.DR2G.Location = new System.Drawing.Point(13, 68);
            this.DR2G.Name = "DR2G";
            this.DR2G.Size = new System.Drawing.Size(95, 17);
            this.DR2G.TabIndex = 3;
            this.DR2G.TabStop = true;
            this.DR2G.Text = "Daily report 2G";
            this.DR2G.UseVisualStyleBackColor = false;
            // 
            // DR3G
            // 
            this.DR3G.AutoSize = true;
            this.DR3G.BackColor = System.Drawing.Color.Transparent;
            this.DR3G.Location = new System.Drawing.Point(13, 91);
            this.DR3G.Name = "DR3G";
            this.DR3G.Size = new System.Drawing.Size(95, 17);
            this.DR3G.TabIndex = 4;
            this.DR3G.TabStop = true;
            this.DR3G.Text = "Daily report 3G";
            this.DR3G.UseVisualStyleBackColor = false;
            // 
            // lst_logfile
            // 
            this.lst_logfile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lst_logfile.FormattingEnabled = true;
            this.lst_logfile.Location = new System.Drawing.Point(0, 0);
            this.lst_logfile.Name = "lst_logfile";
            this.lst_logfile.Size = new System.Drawing.Size(622, 271);
            this.lst_logfile.TabIndex = 17;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.lst_logfile);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(137, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(622, 271);
            this.panel1.TabIndex = 18;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.Transparent;
            this.panel2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("panel2.BackgroundImage")));
            this.panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.Template);
            this.panel2.Controls.Add(this.Run);
            this.panel2.Controls.Add(this.DR3G);
            this.panel2.Controls.Add(this.DR2G);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(137, 271);
            this.panel2.TabIndex = 19;
            // 
            // ToolReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(759, 271);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "ToolReport";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Tool Report V.1";
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button Template;
        private System.Windows.Forms.Button Run;
        private System.Windows.Forms.RadioButton DR2G;
        private System.Windows.Forms.RadioButton DR3G;
        private System.Windows.Forms.ListBox lst_logfile;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
    }
}

