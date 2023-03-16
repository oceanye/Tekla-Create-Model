namespace PlatformCollect
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnDb = new System.Windows.Forms.Button();
            this.btnRevit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnDb
            // 
            this.btnDb.Location = new System.Drawing.Point(174, 119);
            this.btnDb.Name = "btnDb";
            this.btnDb.Size = new System.Drawing.Size(75, 23);
            this.btnDb.TabIndex = 0;
            this.btnDb.Text = "选择数据库";
            this.btnDb.UseVisualStyleBackColor = true;
            this.btnDb.Click += new System.EventHandler(this.btnDb_Click);
            // 
            // btnRevit
            // 
            this.btnRevit.Location = new System.Drawing.Point(174, 245);
            this.btnRevit.Name = "btnRevit";
            this.btnRevit.Size = new System.Drawing.Size(75, 23);
            this.btnRevit.TabIndex = 1;
            this.btnRevit.Text = "Revit";
            this.btnRevit.UseVisualStyleBackColor = true;
            this.btnRevit.Click += new System.EventHandler(this.btnRevit_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(492, 331);
            this.Controls.Add(this.btnRevit);
            this.Controls.Add(this.btnDb);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnDb;
        private System.Windows.Forms.Button btnRevit;
    }
}

