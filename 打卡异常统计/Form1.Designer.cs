namespace 打卡异常统计
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.buttonOpenCheckIn = new System.Windows.Forms.Button();
            this.richTextBoxException = new System.Windows.Forms.RichTextBox();
            this.buttonExit = new System.Windows.Forms.Button();
            this.checkBoxSave = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // buttonOpenCheckIn
            // 
            this.buttonOpenCheckIn.Location = new System.Drawing.Point(528, 88);
            this.buttonOpenCheckIn.Name = "buttonOpenCheckIn";
            this.buttonOpenCheckIn.Size = new System.Drawing.Size(110, 42);
            this.buttonOpenCheckIn.TabIndex = 0;
            this.buttonOpenCheckIn.Text = "导入打卡数据文件";
            this.buttonOpenCheckIn.UseVisualStyleBackColor = true;
            this.buttonOpenCheckIn.Click += new System.EventHandler(this.button1_Click);
            // 
            // richTextBoxException
            // 
            this.richTextBoxException.Location = new System.Drawing.Point(12, 12);
            this.richTextBoxException.Name = "richTextBoxException";
            this.richTextBoxException.ReadOnly = true;
            this.richTextBoxException.Size = new System.Drawing.Size(491, 418);
            this.richTextBoxException.TabIndex = 1;
            this.richTextBoxException.Text = "";
            // 
            // buttonExit
            // 
            this.buttonExit.Location = new System.Drawing.Point(528, 351);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new System.Drawing.Size(110, 61);
            this.buttonExit.TabIndex = 2;
            this.buttonExit.Text = "退出";
            this.buttonExit.UseVisualStyleBackColor = true;
            this.buttonExit.Click += new System.EventHandler(this.buttonExit_Click);
            // 
            // checkBoxSave
            // 
            this.checkBoxSave.AutoSize = true;
            this.checkBoxSave.Checked = true;
            this.checkBoxSave.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSave.Location = new System.Drawing.Point(528, 33);
            this.checkBoxSave.Name = "checkBoxSave";
            this.checkBoxSave.Size = new System.Drawing.Size(90, 16);
            this.checkBoxSave.TabIndex = 3;
            this.checkBoxSave.Text = "保存至EXCLE";
            this.checkBoxSave.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(673, 486);
            this.Controls.Add(this.checkBoxSave);
            this.Controls.Add(this.buttonExit);
            this.Controls.Add(this.richTextBoxException);
            this.Controls.Add(this.buttonOpenCheckIn);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "菜菜的打卡统计工具 V1.2.0";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonOpenCheckIn;
        private System.Windows.Forms.RichTextBox richTextBoxException;
        private System.Windows.Forms.Button buttonExit;
        private System.Windows.Forms.CheckBox checkBoxSave;
    }
}

