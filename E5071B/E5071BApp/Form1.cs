using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using E5071BLib;

namespace E5071BApp
{
    public partial class Form1 : Form
    {
        E5071BApi m_ena;
        int countShots = 0;
        public Form1()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;

            E5071BLib.E5071BApi.VNACallback p = new E5071BApi.VNACallback(VNACallbackMesages);
            m_ena = new E5071BApi(p);
        }

        void VNACallbackMesages(int code , string msg)
        {
            switch (code)
            {
                case 200:
                    label5.Text = msg;
                    countShots++;
                    label6.Text = countShots.ToString();
                break;
                case 868:
                    label4.Text = msg;
                break;
                case 860:
                    label7.Text = "Left in queue: " + msg;
                break;

            }
        }
        private void button1_Click(object sender, EventArgs e)
        {            
            string STACalFile = "d:\\homertuner.sta";
            try
            {
                //m_ena.Initialize(textBox1.Text,  STACalFile);
                m_ena.Initialize(textBox1.Text, "y:\\" , STACalFile);
               // m_ena.Start(@"C:\mpfmtests", "y:\\");
            }
            catch (Exception  err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            m_ena.SingleShotCapture("test.s2p");

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_ena.Stop();
            m_ena.Close();
        }
    }
}
