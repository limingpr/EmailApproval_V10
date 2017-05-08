namespace EmailApproval_V10
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
            this.components = new System.ComponentModel.Container();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.txtWorkTime = new System.Windows.Forms.TextBox();
            this.txtFormTime = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnEnd = new System.Windows.Forms.Button();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.label4 = new System.Windows.Forms.Label();
            this.txtLastTime = new System.Windows.Forms.TextBox();
            this.timer3 = new System.Windows.Forms.Timer(this.components);
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnTest = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // timer2
            // 
            this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(100, 94);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(226, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "轮训待办时间（秒）";
            // 
            // txtWorkTime
            // 
            this.txtWorkTime.Location = new System.Drawing.Point(356, 88);
            this.txtWorkTime.Margin = new System.Windows.Forms.Padding(6);
            this.txtWorkTime.Name = "txtWorkTime";
            this.txtWorkTime.Size = new System.Drawing.Size(196, 35);
            this.txtWorkTime.TabIndex = 1;
            this.txtWorkTime.Text = "3";
            // 
            // txtFormTime
            // 
            this.txtFormTime.Location = new System.Drawing.Point(356, 174);
            this.txtFormTime.Margin = new System.Windows.Forms.Padding(6);
            this.txtFormTime.Name = "txtFormTime";
            this.txtFormTime.Size = new System.Drawing.Size(196, 35);
            this.txtFormTime.TabIndex = 2;
            this.txtFormTime.Text = "10";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(100, 192);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(226, 24);
            this.label2.TabIndex = 3;
            this.label2.Text = "表单加载时间（秒）";
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(104, 292);
            this.btnStart.Margin = new System.Windows.Forms.Padding(6);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(150, 46);
            this.btnStart.TabIndex = 4;
            this.btnStart.Text = "开始轮训";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.BtnStart_Click);
            // 
            // btnEnd
            // 
            this.btnEnd.Location = new System.Drawing.Point(356, 292);
            this.btnEnd.Margin = new System.Windows.Forms.Padding(6);
            this.btnEnd.Name = "btnEnd";
            this.btnEnd.Size = new System.Drawing.Size(150, 46);
            this.btnEnd.TabIndex = 5;
            this.btnEnd.Text = "结束轮训";
            this.btnEnd.UseVisualStyleBackColor = true;
            this.btnEnd.Click += new System.EventHandler(this.button2_Click);
            // 
            // webBrowser1
            // 
            this.webBrowser1.Location = new System.Drawing.Point(86, 378);
            this.webBrowser1.Margin = new System.Windows.Forms.Padding(6);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(40, 40);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(1698, 682);
            this.webBrowser1.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(100, 18);
            this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(226, 24);
            this.label4.TabIndex = 8;
            this.label4.Text = "上次结束时间（分）";
            // 
            // txtLastTime
            // 
            this.txtLastTime.Location = new System.Drawing.Point(356, 18);
            this.txtLastTime.Margin = new System.Windows.Forms.Padding(6);
            this.txtLastTime.Name = "txtLastTime";
            this.txtLastTime.Size = new System.Drawing.Size(196, 35);
            this.txtLastTime.TabIndex = 9;
            this.txtLastTime.Text = "60";
            // 
            // timer3
            // 
            this.timer3.Interval = 2000;
            this.timer3.Tick += new System.EventHandler(this.timer3_Tick);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(819, 94);
            this.textBox1.Margin = new System.Windows.Forms.Padding(6);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(474, 35);
            this.textBox1.TabIndex = 10;
            this.textBox1.Text = "dd17c0c0-1b5b-44d3-902d-fcc2e97a2248";
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(819, 292);
            this.btnTest.Margin = new System.Windows.Forms.Padding(6);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(150, 46);
            this.btnTest.TabIndex = 11;
            this.btnTest.Text = "测试转换";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(815, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(178, 24);
            this.label3.TabIndex = 12;
            this.label3.Text = "测试WorkItemID";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1862, 1120);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.txtLastTime);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.btnEnd);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtFormTime);
            this.Controls.Add(this.txtWorkTime);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Timer timer2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtWorkTime;
        private System.Windows.Forms.TextBox txtFormTime;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnEnd;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtLastTime;
        private System.Windows.Forms.Timer timer3;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.Label label3;
    }
}

