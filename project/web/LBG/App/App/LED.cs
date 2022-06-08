using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace App
{
    public partial class LED : System.Windows.Forms.Panel
    {
        private bool status;
        public LED()
        {
            InitializeComponent();
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
        }

        public bool on
        {
            get
            {
                return status;
            }
            set
            {
                status = value;
                if (status) this.BackColor = System.Drawing.Color.Red;
                else
                    this.BackColor = System.Drawing.Color.Transparent;

                this.Refresh();
                System.Threading.Thread.Sleep(100);
            }
        }
    }
}
