namespace TestTekla
{
    partial class ShowForm
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
            this.listB = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // listB
            // 
            this.listB.FormattingEnabled = true;
            this.listB.ItemHeight = 12;
            this.listB.Location = new System.Drawing.Point(1, -3);
            this.listB.Name = "listB";
            this.listB.Size = new System.Drawing.Size(313, 268);
            this.listB.TabIndex = 0;
            // 
            // ShowForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(314, 261);
            this.Controls.Add(this.listB);
            this.Name = "ShowForm";
            this.Text = "ShowForm";
            this.Load += new System.EventHandler(this.ShowForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox listB;
    }
}