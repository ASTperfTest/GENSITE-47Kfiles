namespace App
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.btn_Start = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.nmTimerExecute = new System.Windows.Forms.NumericUpDown();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.nm1Hrs = new System.Windows.Forms.NumericUpDown();
            this.btn_Stop = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkedListBox2 = new System.Windows.Forms.CheckedListBox();
            this.lbl1Days = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.checkedListBox3 = new System.Windows.Forms.CheckedListBox();
            this.lbl5Days = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.numRankingInterval = new System.Windows.Forms.NumericUpDown();
            this.checkedListBox4 = new System.Windows.Forms.CheckedListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.txtConnstr = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.LED_TIMER2 = new App.LED();
            this.LED_TIMER1 = new App.LED();
            this.led_SP_SCORE_RANKING = new App.LED();
            this.led_SP_STARS = new App.LED();
            this.led_SP_SCORE = new App.LED();
            this.led_SP_BUG = new App.LED();
            this.led_SP_WEATHER = new App.LED();
            this.led_SP_FERTY_DECAY = new App.LED();
            this.led_SP_WATER_DECAY = new App.LED();
            this.led_SP_CHECKDEAD = new App.LED();
            this.led_SP_CLOCK = new App.LED();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nmTimerExecute)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nm1Hrs)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numRankingInterval)).BeginInit();
            this.SuspendLayout();
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // btn_Start
            // 
            this.btn_Start.Location = new System.Drawing.Point(444, 12);
            this.btn_Start.Name = "btn_Start";
            this.btn_Start.Size = new System.Drawing.Size(75, 23);
            this.btn_Start.TabIndex = 0;
            this.btn_Start.Text = "啟動";
            this.btn_Start.UseVisualStyleBackColor = true;
            this.btn_Start.Click += new System.EventHandler(this.btn_Start_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.led_SP_CHECKDEAD);
            this.groupBox1.Controls.Add(this.led_SP_CLOCK);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.nmTimerExecute);
            this.groupBox1.Controls.Add(this.checkedListBox1);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.nm1Hrs);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(421, 129);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "每1遊戲小時執行工作";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(395, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(17, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "秒";
            // 
            // nmTimerExecute
            // 
            this.nmTimerExecute.Location = new System.Drawing.Point(309, 21);
            this.nmTimerExecute.Maximum = new decimal(new int[] {
            60,
            0,
            0,
            0});
            this.nmTimerExecute.Minimum = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.nmTimerExecute.Name = "nmTimerExecute";
            this.nmTimerExecute.Size = new System.Drawing.Size(80, 22);
            this.nmTimerExecute.TabIndex = 3;
            this.nmTimerExecute.Value = new decimal(new int[] {
            60,
            0,
            0,
            0});
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Items.AddRange(new object[] {
            "SP_CLOCK -- 變換時鐘",
            "SP_CHECKDEAD --> 檢查植物狀態"});
            this.checkedListBox1.Location = new System.Drawing.Point(6, 49);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(383, 72);
            this.checkedListBox1.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(132, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "秒";
            // 
            // nm1Hrs
            // 
            this.nm1Hrs.Location = new System.Drawing.Point(6, 21);
            this.nm1Hrs.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.nm1Hrs.Minimum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.nm1Hrs.Name = "nm1Hrs";
            this.nm1Hrs.Size = new System.Drawing.Size(120, 22);
            this.nm1Hrs.TabIndex = 0;
            this.nm1Hrs.Value = new decimal(new int[] {
            3600,
            0,
            0,
            0});
            this.nm1Hrs.ValueChanged += new System.EventHandler(this.nm1Hrs_ValueChanged);
            // 
            // btn_Stop
            // 
            this.btn_Stop.Enabled = false;
            this.btn_Stop.Location = new System.Drawing.Point(444, 48);
            this.btn_Stop.Name = "btn_Stop";
            this.btn_Stop.Size = new System.Drawing.Size(75, 23);
            this.btn_Stop.TabIndex = 2;
            this.btn_Stop.Text = "暫停";
            this.btn_Stop.UseVisualStyleBackColor = true;
            this.btn_Stop.Click += new System.EventHandler(this.btn_Stop_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.led_SP_WEATHER);
            this.groupBox2.Controls.Add(this.led_SP_FERTY_DECAY);
            this.groupBox2.Controls.Add(this.led_SP_WATER_DECAY);
            this.groupBox2.Controls.Add(this.checkedListBox2);
            this.groupBox2.Controls.Add(this.lbl1Days);
            this.groupBox2.Location = new System.Drawing.Point(12, 147);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(421, 130);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "每1遊戲日執行工作";
            // 
            // checkedListBox2
            // 
            this.checkedListBox2.FormattingEnabled = true;
            this.checkedListBox2.Items.AddRange(new object[] {
            "SP_WATER_DECAY --> 水分度減少",
            "SP_FERTY_DECAY --> 肥料度減少",
            "SP_WEATHER  --> 天氣變換 ",
            "更新 LAST_TXN_TIMESTAMP"});
            this.checkedListBox2.Location = new System.Drawing.Point(6, 33);
            this.checkedListBox2.Name = "checkedListBox2";
            this.checkedListBox2.Size = new System.Drawing.Size(383, 89);
            this.checkedListBox2.TabIndex = 5;
            // 
            // lbl1Days
            // 
            this.lbl1Days.AutoSize = true;
            this.lbl1Days.Location = new System.Drawing.Point(4, 18);
            this.lbl1Days.Name = "lbl1Days";
            this.lbl1Days.Size = new System.Drawing.Size(44, 12);
            this.lbl1Days.TabIndex = 4;
            this.lbl1Days.Text = "8640 秒";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.led_SP_STARS);
            this.groupBox3.Controls.Add(this.led_SP_SCORE);
            this.groupBox3.Controls.Add(this.led_SP_BUG);
            this.groupBox3.Controls.Add(this.checkedListBox3);
            this.groupBox3.Controls.Add(this.lbl5Days);
            this.groupBox3.Location = new System.Drawing.Point(13, 283);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(420, 129);
            this.groupBox3.TabIndex = 4;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "每5遊戲日執行工作";
            // 
            // checkedListBox3
            // 
            this.checkedListBox3.FormattingEnabled = true;
            this.checkedListBox3.Items.AddRange(new object[] {
            "SP_BUG-->  蟲害度評估",
            "SP_SCORE -->  評分",
            "SP_STARS--> 評星等      ",
            "更新 LAST_CHECK_STATE"});
            this.checkedListBox3.Location = new System.Drawing.Point(5, 30);
            this.checkedListBox3.Name = "checkedListBox3";
            this.checkedListBox3.Size = new System.Drawing.Size(383, 89);
            this.checkedListBox3.TabIndex = 8;
            // 
            // lbl5Days
            // 
            this.lbl5Days.AutoSize = true;
            this.lbl5Days.Location = new System.Drawing.Point(3, 18);
            this.lbl5Days.Name = "lbl5Days";
            this.lbl5Days.Size = new System.Drawing.Size(50, 12);
            this.lbl5Days.TabIndex = 7;
            this.lbl5Days.Text = "43200 秒";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.led_SP_SCORE_RANKING);
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.numRankingInterval);
            this.groupBox4.Controls.Add(this.checkedListBox4);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Location = new System.Drawing.Point(12, 418);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(420, 65);
            this.groupBox4.TabIndex = 5;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "遊戲總分匯整";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(395, 20);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(17, 12);
            this.label4.TabIndex = 10;
            this.label4.Text = "秒";
            // 
            // numRankingInterval
            // 
            this.numRankingInterval.Location = new System.Drawing.Point(309, 10);
            this.numRankingInterval.Maximum = new decimal(new int[] {
            86400,
            0,
            0,
            0});
            this.numRankingInterval.Minimum = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.numRankingInterval.Name = "numRankingInterval";
            this.numRankingInterval.Size = new System.Drawing.Size(80, 22);
            this.numRankingInterval.TabIndex = 9;
            this.numRankingInterval.Value = new decimal(new int[] {
            3600,
            0,
            0,
            0});
            // 
            // checkedListBox4
            // 
            this.checkedListBox4.FormattingEnabled = true;
            this.checkedListBox4.Items.AddRange(new object[] {
            "SP_SCORE_RANKING --> 總分計算"});
            this.checkedListBox4.Location = new System.Drawing.Point(6, 38);
            this.checkedListBox4.Name = "checkedListBox4";
            this.checkedListBox4.Size = new System.Drawing.Size(383, 21);
            this.checkedListBox4.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 12);
            this.label3.TabIndex = 7;
            this.label3.Text = "1hrs (3600 秒)";
            // 
            // timer2
            // 
            this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
            // 
            // txtConnstr
            // 
            this.txtConnstr.Location = new System.Drawing.Point(13, 511);
            this.txtConnstr.Name = "txtConnstr";
            this.txtConnstr.Size = new System.Drawing.Size(420, 22);
            this.txtConnstr.TabIndex = 10;
            this.txtConnstr.Text = "Data Source=localhost;Initial Catalog=LITTLE_BEAN;Persist Security Info=True;User" +
                " ID=sa;Password=itc";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 490);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(90, 12);
            this.label5.TabIndex = 11;
            this.label5.Text = "Connection String";
            // 
            // LED_TIMER2
            // 
            this.LED_TIMER2.BackColor = System.Drawing.Color.Transparent;
            this.LED_TIMER2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.LED_TIMER2.Location = new System.Drawing.Point(446, 101);
            this.LED_TIMER2.Name = "LED_TIMER2";
            this.LED_TIMER2.on = false;
            this.LED_TIMER2.Size = new System.Drawing.Size(24, 18);
            this.LED_TIMER2.TabIndex = 9;
            // 
            // LED_TIMER1
            // 
            this.LED_TIMER1.BackColor = System.Drawing.Color.Transparent;
            this.LED_TIMER1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.LED_TIMER1.Location = new System.Drawing.Point(446, 77);
            this.LED_TIMER1.Name = "LED_TIMER1";
            this.LED_TIMER1.on = false;
            this.LED_TIMER1.Size = new System.Drawing.Size(24, 18);
            this.LED_TIMER1.TabIndex = 8;
            // 
            // led_SP_SCORE_RANKING
            // 
            this.led_SP_SCORE_RANKING.BackColor = System.Drawing.Color.Transparent;
            this.led_SP_SCORE_RANKING.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.led_SP_SCORE_RANKING.Location = new System.Drawing.Point(395, 40);
            this.led_SP_SCORE_RANKING.Name = "led_SP_SCORE_RANKING";
            this.led_SP_SCORE_RANKING.on = false;
            this.led_SP_SCORE_RANKING.Size = new System.Drawing.Size(15, 16);
            this.led_SP_SCORE_RANKING.TabIndex = 12;
            // 
            // led_SP_STARS
            // 
            this.led_SP_STARS.BackColor = System.Drawing.Color.Transparent;
            this.led_SP_STARS.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.led_SP_STARS.Location = new System.Drawing.Point(394, 74);
            this.led_SP_STARS.Name = "led_SP_STARS";
            this.led_SP_STARS.on = false;
            this.led_SP_STARS.Size = new System.Drawing.Size(15, 16);
            this.led_SP_STARS.TabIndex = 11;
            // 
            // led_SP_SCORE
            // 
            this.led_SP_SCORE.BackColor = System.Drawing.Color.Transparent;
            this.led_SP_SCORE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.led_SP_SCORE.Location = new System.Drawing.Point(394, 52);
            this.led_SP_SCORE.Name = "led_SP_SCORE";
            this.led_SP_SCORE.on = false;
            this.led_SP_SCORE.Size = new System.Drawing.Size(15, 16);
            this.led_SP_SCORE.TabIndex = 11;
            // 
            // led_SP_BUG
            // 
            this.led_SP_BUG.BackColor = System.Drawing.Color.Transparent;
            this.led_SP_BUG.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.led_SP_BUG.Location = new System.Drawing.Point(394, 30);
            this.led_SP_BUG.Name = "led_SP_BUG";
            this.led_SP_BUG.on = false;
            this.led_SP_BUG.Size = new System.Drawing.Size(15, 16);
            this.led_SP_BUG.TabIndex = 11;
            // 
            // led_SP_WEATHER
            // 
            this.led_SP_WEATHER.BackColor = System.Drawing.Color.Transparent;
            this.led_SP_WEATHER.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.led_SP_WEATHER.Location = new System.Drawing.Point(395, 77);
            this.led_SP_WEATHER.Name = "led_SP_WEATHER";
            this.led_SP_WEATHER.on = false;
            this.led_SP_WEATHER.Size = new System.Drawing.Size(15, 16);
            this.led_SP_WEATHER.TabIndex = 11;
            // 
            // led_SP_FERTY_DECAY
            // 
            this.led_SP_FERTY_DECAY.BackColor = System.Drawing.Color.Transparent;
            this.led_SP_FERTY_DECAY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.led_SP_FERTY_DECAY.Location = new System.Drawing.Point(395, 55);
            this.led_SP_FERTY_DECAY.Name = "led_SP_FERTY_DECAY";
            this.led_SP_FERTY_DECAY.on = false;
            this.led_SP_FERTY_DECAY.Size = new System.Drawing.Size(15, 16);
            this.led_SP_FERTY_DECAY.TabIndex = 12;
            // 
            // led_SP_WATER_DECAY
            // 
            this.led_SP_WATER_DECAY.BackColor = System.Drawing.Color.Transparent;
            this.led_SP_WATER_DECAY.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.led_SP_WATER_DECAY.Location = new System.Drawing.Point(395, 33);
            this.led_SP_WATER_DECAY.Name = "led_SP_WATER_DECAY";
            this.led_SP_WATER_DECAY.on = false;
            this.led_SP_WATER_DECAY.Size = new System.Drawing.Size(15, 16);
            this.led_SP_WATER_DECAY.TabIndex = 11;
            // 
            // led_SP_CHECKDEAD
            // 
            this.led_SP_CHECKDEAD.BackColor = System.Drawing.Color.Transparent;
            this.led_SP_CHECKDEAD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.led_SP_CHECKDEAD.Location = new System.Drawing.Point(395, 71);
            this.led_SP_CHECKDEAD.Name = "led_SP_CHECKDEAD";
            this.led_SP_CHECKDEAD.on = false;
            this.led_SP_CHECKDEAD.Size = new System.Drawing.Size(15, 16);
            this.led_SP_CHECKDEAD.TabIndex = 11;
            // 
            // led_SP_CLOCK
            // 
            this.led_SP_CLOCK.BackColor = System.Drawing.Color.Transparent;
            this.led_SP_CLOCK.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.led_SP_CLOCK.Location = new System.Drawing.Point(395, 49);
            this.led_SP_CLOCK.Name = "led_SP_CLOCK";
            this.led_SP_CLOCK.on = false;
            this.led_SP_CLOCK.Size = new System.Drawing.Size(15, 16);
            this.led_SP_CLOCK.TabIndex = 10;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(531, 545);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtConnstr);
            this.Controls.Add(this.LED_TIMER2);
            this.Controls.Add(this.LED_TIMER1);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btn_Stop);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btn_Start);
            this.Name = "Form1";
            this.Text = "豆仔世界 -- 時間驅動器";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nmTimerExecute)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nm1Hrs)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numRankingInterval)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button btn_Start;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown nm1Hrs;
        private System.Windows.Forms.Button btn_Stop;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.CheckedListBox checkedListBox2;
        private System.Windows.Forms.Label lbl1Days;
        private System.Windows.Forms.CheckedListBox checkedListBox3;
        private System.Windows.Forms.Label lbl5Days;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown nmTimerExecute;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.CheckedListBox checkedListBox4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown numRankingInterval;
        private System.Windows.Forms.Timer timer2;
        private LED LED_TIMER1;
        private LED LED_TIMER2;
        private LED led_SP_CLOCK;
        private LED led_SP_WATER_DECAY;
        private LED led_SP_WEATHER;
        private LED led_SP_FERTY_DECAY;
        private LED led_SP_STARS;
        private LED led_SP_SCORE;
        private LED led_SP_BUG;
        private LED led_SP_SCORE_RANKING;
        private LED led_SP_CHECKDEAD;
        private System.Windows.Forms.TextBox txtConnstr;
        private System.Windows.Forms.Label label5;
    }
}

