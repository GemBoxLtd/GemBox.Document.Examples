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
            this.MenuStrip = new System.Windows.Forms.MenuStrip();
            this.LoadFileMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.PrintFileMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.PrintPreviewControl = new System.Windows.Forms.PrintPreviewControl();
            this.PageLb = new System.Windows.Forms.Label();
            this.PageUpDown = new System.Windows.Forms.NumericUpDown();
            this.MenuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PageUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // MenuStrip
            // 
            this.MenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.LoadFileMenuItem,
            this.PrintFileMenuItem});
            this.MenuStrip.Location = new System.Drawing.Point(0, 0);
            this.MenuStrip.Name = "MenuStrip";
            this.MenuStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
            this.MenuStrip.Size = new System.Drawing.Size(784, 24);
            this.MenuStrip.TabIndex = 0;
            this.MenuStrip.Text = "MenuStrip";
            // 
            // LoadFileMenuItem
            // 
            this.LoadFileMenuItem.Name = "LoadFileMenuItem";
            this.LoadFileMenuItem.Size = new System.Drawing.Size(45, 20);
            this.LoadFileMenuItem.Text = "Load";
            this.LoadFileMenuItem.Click += new System.EventHandler(this.LoadFileMenuItem_Click);
            // 
            // PrintFileMenuItem
            // 
            this.PrintFileMenuItem.Name = "PrintFileMenuItem";
            this.PrintFileMenuItem.Size = new System.Drawing.Size(44, 20);
            this.PrintFileMenuItem.Text = "Print";
            this.PrintFileMenuItem.Click += new System.EventHandler(this.PrintFileMenuItem_Click);
            // 
            // PrintPreviewControl
            // 
            this.PrintPreviewControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.PrintPreviewControl.Location = new System.Drawing.Point(0, 53);
            this.PrintPreviewControl.Name = "PrintPreviewControl";
            this.PrintPreviewControl.Size = new System.Drawing.Size(784, 358);
            this.PrintPreviewControl.TabIndex = 1;
            // 
            // PageLb
            // 
            this.PageLb.AutoSize = true;
            this.PageLb.Location = new System.Drawing.Point(12, 33);
            this.PageLb.Name = "PageLb";
            this.PageLb.Size = new System.Drawing.Size(35, 13);
            this.PageLb.TabIndex = 2;
            this.PageLb.Text = "Page:";
            // 
            // PageUpDown
            // 
            this.PageUpDown.Location = new System.Drawing.Point(53, 29);
            this.PageUpDown.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.PageUpDown.Name = "PageUpDown";
            this.PageUpDown.Size = new System.Drawing.Size(40, 20);
            this.PageUpDown.TabIndex = 3;
            this.PageUpDown.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.PageUpDown.ValueChanged += new System.EventHandler(this.PageUpDown_ValueChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 411);
            this.Controls.Add(this.PageUpDown);
            this.Controls.Add(this.PageLb);
            this.Controls.Add(this.PrintPreviewControl);
            this.Controls.Add(this.MenuStrip);
            this.Name = "Form1";
            this.Text = "Printing in Windows Forms application";
            this.MenuStrip.ResumeLayout(false);
            this.MenuStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PageUpDown)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.MenuStrip MenuStrip;
    private System.Windows.Forms.ToolStripMenuItem LoadFileMenuItem;
    private System.Windows.Forms.ToolStripMenuItem PrintFileMenuItem;
    private System.Windows.Forms.PrintPreviewControl PrintPreviewControl;
    private System.Windows.Forms.Label PageLb;
    private System.Windows.Forms.NumericUpDown PageUpDown;
}
