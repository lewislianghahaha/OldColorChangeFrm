namespace OldColorChangeFrm
{
    partial class Main
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
            this.MainMenu = new System.Windows.Forms.MenuStrip();
            this.tmclose = new System.Windows.Forms.ToolStripMenuItem();
            this.btngen = new System.Windows.Forms.Button();
            this.MainMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // MainMenu
            // 
            this.MainMenu.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.MainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tmclose});
            this.MainMenu.Location = new System.Drawing.Point(0, 0);
            this.MainMenu.Name = "MainMenu";
            this.MainMenu.Size = new System.Drawing.Size(284, 25);
            this.MainMenu.TabIndex = 0;
            this.MainMenu.Text = "menuStrip1";
            // 
            // tmclose
            // 
            this.tmclose.Name = "tmclose";
            this.tmclose.Size = new System.Drawing.Size(44, 21);
            this.tmclose.Text = "关闭";
            // 
            // btngen
            // 
            this.btngen.Location = new System.Drawing.Point(48, 38);
            this.btngen.Name = "btngen";
            this.btngen.Size = new System.Drawing.Size(175, 23);
            this.btngen.TabIndex = 1;
            this.btngen.TabStop = false;
            this.btngen.Text = "运算及调出";
            this.btngen.UseVisualStyleBackColor = true;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 82);
            this.ControlBox = false;
            this.Controls.Add(this.btngen);
            this.Controls.Add(this.MainMenu);
            this.MainMenuStrip = this.MainMenu;
            this.Name = "Main";
            this.Text = "旧系统配方调出工具";
            this.MainMenu.ResumeLayout(false);
            this.MainMenu.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip MainMenu;
        private System.Windows.Forms.ToolStripMenuItem tmclose;
        private System.Windows.Forms.Button btngen;
    }
}

