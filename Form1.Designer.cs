﻿namespace tx_custom_fields_SUM
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
            this.components = new System.ComponentModel.Container();
            this.textControl1 = new TXTextControl.TextControl();
            this.buttonBar1 = new TXTextControl.ButtonBar();
            this.rulerBar2 = new TXTextControl.RulerBar();
            this.statusBar1 = new TXTextControl.StatusBar();
            this.rulerBar1 = new TXTextControl.RulerBar();
            this.mailMerge1 = new TXTextControl.DocumentServer.MailMerge(this.components);
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.mailMergeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadTemplateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mergeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // textControl1
            // 
            this.textControl1.ButtonBar = this.buttonBar1;
            this.textControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textControl1.Font = new System.Drawing.Font("Arial", 10F);
            this.textControl1.Location = new System.Drawing.Point(25, 77);
            this.textControl1.Name = "textControl1";
            this.textControl1.RulerBar = this.rulerBar2;
            this.textControl1.Size = new System.Drawing.Size(869, 415);
            this.textControl1.StatusBar = this.statusBar1;
            this.textControl1.TabIndex = 0;
            this.textControl1.VerticalRulerBar = this.rulerBar1;
            // 
            // buttonBar1
            // 
            this.buttonBar1.BackColor = System.Drawing.SystemColors.Control;
            this.buttonBar1.Dock = System.Windows.Forms.DockStyle.Top;
            this.buttonBar1.Location = new System.Drawing.Point(0, 24);
            this.buttonBar1.Name = "buttonBar1";
            this.buttonBar1.Size = new System.Drawing.Size(894, 28);
            this.buttonBar1.TabIndex = 1;
            this.buttonBar1.Text = "buttonBar1";
            // 
            // rulerBar2
            // 
            this.rulerBar2.Dock = System.Windows.Forms.DockStyle.Top;
            this.rulerBar2.Location = new System.Drawing.Point(0, 52);
            this.rulerBar2.Name = "rulerBar2";
            this.rulerBar2.Size = new System.Drawing.Size(894, 25);
            this.rulerBar2.TabIndex = 4;
            this.rulerBar2.Text = "rulerBar2";
            // 
            // statusBar1
            // 
            this.statusBar1.BackColor = System.Drawing.SystemColors.Control;
            this.statusBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.statusBar1.Location = new System.Drawing.Point(0, 492);
            this.statusBar1.Name = "statusBar1";
            this.statusBar1.Size = new System.Drawing.Size(894, 22);
            this.statusBar1.TabIndex = 2;
            // 
            // rulerBar1
            // 
            this.rulerBar1.Alignment = TXTextControl.RulerBarAlignment.Left;
            this.rulerBar1.Dock = System.Windows.Forms.DockStyle.Left;
            this.rulerBar1.Location = new System.Drawing.Point(0, 77);
            this.rulerBar1.Name = "rulerBar1";
            this.rulerBar1.Size = new System.Drawing.Size(25, 415);
            this.rulerBar1.TabIndex = 3;
            this.rulerBar1.Text = "rulerBar1";
            // 
            // mailMerge1
            // 
            this.mailMerge1.RemoveEmptyFields = false;
            this.mailMerge1.ReportDataSourceConfig = "";
            this.mailMerge1.TemplateFile = "template.docx";
            this.mailMerge1.TextComponent = this.textControl1;
            this.mailMerge1.DataRowMerged += new TXTextControl.DocumentServer.MailMerge.DataRowMergedHandler(this.mailMerge1_DataRowMerged);
            this.mailMerge1.FieldMerged += new TXTextControl.DocumentServer.MailMerge.FieldMergedHandler(this.mailMerge1_FieldMerged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mailMergeToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(894, 24);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // mailMergeToolStripMenuItem
            // 
            this.mailMergeToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.loadTemplateToolStripMenuItem,
            this.mergeToolStripMenuItem});
            this.mailMergeToolStripMenuItem.Name = "mailMergeToolStripMenuItem";
            this.mailMergeToolStripMenuItem.Size = new System.Drawing.Size(76, 20);
            this.mailMergeToolStripMenuItem.Text = "MailMerge";
            this.mailMergeToolStripMenuItem.DropDownOpening += new System.EventHandler(this.mailMergeToolStripMenuItem_DropDownOpening);
            // 
            // loadTemplateToolStripMenuItem
            // 
            this.loadTemplateToolStripMenuItem.Name = "loadTemplateToolStripMenuItem";
            this.loadTemplateToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.loadTemplateToolStripMenuItem.Text = "Load Template";
            this.loadTemplateToolStripMenuItem.Click += new System.EventHandler(this.loadTemplateToolStripMenuItem_Click);
            // 
            // mergeToolStripMenuItem
            // 
            this.mergeToolStripMenuItem.Name = "mergeToolStripMenuItem";
            this.mergeToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.mergeToolStripMenuItem.Text = "Merge";
            this.mergeToolStripMenuItem.Click += new System.EventHandler(this.mergeToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(894, 514);
            this.Controls.Add(this.textControl1);
            this.Controls.Add(this.rulerBar1);
            this.Controls.Add(this.rulerBar2);
            this.Controls.Add(this.buttonBar1);
            this.Controls.Add(this.statusBar1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Text Control Reporting: Custom Field Processing";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private TXTextControl.TextControl textControl1;
        private TXTextControl.ButtonBar buttonBar1;
        private TXTextControl.RulerBar rulerBar2;
        private TXTextControl.StatusBar statusBar1;
        private TXTextControl.RulerBar rulerBar1;
        private TXTextControl.DocumentServer.MailMerge mailMerge1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem mailMergeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadTemplateToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mergeToolStripMenuItem;
    }
}

