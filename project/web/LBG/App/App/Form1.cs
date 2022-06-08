using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using TimeCore;

namespace App
{
    public partial class Form1 : Form
    {

        private TimeCore.TimeCore core= new TimeCore.TimeCore("Data Source=localhost;Initial Catalog=LITTLE_BEAN;Persist Security Info=True;User ID=sa;Password=itc");
        public Form1()
        {
            InitializeComponent();
            core.accounce += new TimeCore.TimeCore.timerEvent(core_accounce);
        }

        void core_accounce(string eventType, int executingStatus)
        {
            if (eventType == "SP_WEAHTER") led_SP_WEATHER.on = (executingStatus == 1);
            if (eventType == "SP_CLOCK") this.led_SP_CLOCK .on = (executingStatus == 1);
            if (eventType == "SP_SCORE") led_SP_SCORE.on = (executingStatus == 1);
            if (eventType == "SP_STAR") led_SP_STARS.on = (executingStatus == 1);
            if (eventType == "SP_SCORE_RANKING") led_SP_SCORE_RANKING.on = (executingStatus == 1);
            if (eventType == "SP_FERTITY_DECAY") this.led_SP_FERTY_DECAY.on = (executingStatus == 1);
            if (eventType == "SP_WATER_DECAY") this.led_SP_WATER_DECAY.on = (executingStatus == 1);

            if (eventType == "SP_CHECK_DEAD") this.led_SP_CHECKDEAD.on = (executingStatus == 1);

            
        }

        private void nm1Hrs_ValueChanged(object sender, EventArgs e)
        {
            this.lbl1Days.Text  = String.Format("{0} 秒", nm1Hrs.Value * 24); // 24hrs = 1日
            this.lbl5Days.Text = String.Format("{0} 秒", nm1Hrs.Value * 24*5);// 5日
        }

        private void btn_Start_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            timer2.Enabled = false;
            txtConnstr.Enabled = false;
            core.connStr = txtConnstr.Text;
            core.hours = (int)nm1Hrs.Value;
            timer1.Interval = (int)(nmTimerExecute.Value) * 1000;
            timer2.Interval = (int)(numRankingInterval.Value )*1000;
            btn_Stop.Enabled = true;
            btn_Start.Enabled = false;
            timer1.Enabled = true;
            timer2.Enabled = true;
        }

        private void btn_Stop_Click(object sender, EventArgs e)
        {
            txtConnstr.Enabled = true;
            btn_Stop.Enabled = false;
            btn_Start.Enabled = true;
            timer1.Enabled = false;
            timer2.Enabled = false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            LED_TIMER1.on = true;
            core.OneTick();
            LED_TIMER1.on = false;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {

            LED_TIMER2.on = true;
            core.ranking();
            LED_TIMER2.on = false;
        }
    }
}
