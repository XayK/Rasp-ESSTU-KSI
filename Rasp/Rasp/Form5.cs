using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;

namespace SheduleSI
{
    public partial class Form5 : Form
    {
        System.Timers.Timer myTimer = new System.Timers.Timer();
        public Form5()
        {
            InitializeComponent();
            myTimer.Elapsed += new ElapsedEventHandler(DisplayTimeEvent);
            myTimer.Interval = 50; // 1000 ms is one second
            myTimer.Start();
        }

        public void draws()
        {
            myTimer.Start();
        }
        

        public void DisplayTimeEvent(object source, ElapsedEventArgs e)
        {
            pictureBox1.Invalidate();
            
            if (!Run)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    this.Hide();
                });
                myTimer.Stop();
            }
        }
        public bool Run = true;
       
        float start = 0;
        float end = 0;

        private void pictureBox1_Paint_1(object sender, PaintEventArgs e)
        {
            Pen myPen = new Pen(Brushes.Aquamarine);
            myPen.Width = 3.5F;
            float w = pictureBox1.Width;
            float h = pictureBox1.Height;
            e.Graphics.DrawArc(myPen, 0 , 0, w, h, start, end);
            start -= 4;
            end -= 4;
            if (start < -360) start += 360;
            if (end < -360) end += 360;


        }
    }
}
