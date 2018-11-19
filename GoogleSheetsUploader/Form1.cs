﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GoogleSheetsUploader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        bool processing = false;
        int runs = 0;

        private void Form1_Load(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;

            int timer = int.Parse(System.Configuration.ConfigurationManager.AppSettings["timer"]);

            if (timer > 0)
            {
                timer1.Interval = timer * 1000;
                timer1.Start();
                label1.Text = "Timer every " + timer + "s";
            }
            else
                label1.Text = "No timer";
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized )
            {
                notifyIcon1.Visible = true;
                Hide();
            }
        }

        private void Send_Click(object sender, EventArgs e)
        {
            if (!processing)
                backgroundWorker1.RunWorkerAsync();
        }

        private void Process()
        {
            if (processing)
                return;

            try
            {
                processing = true;

                label1.Text = "Start: " + DateTime.Now;
                textBox1.Text = "";
                LogProcessor logProcessor = new LogProcessor();
                logProcessor.OnMessage += LogProcessor_OnMessage;
                logProcessor.Process();

                runs++;

                label1.Text += "... End: " + DateTime.Now + " (" + runs + ")";
                notifyIcon1.Text = label1.Text;
            }
            finally
            {
                processing = false;
            }
        }

        private void LogProcessor_OnMessage(string msg)
        {
            textBox1.Text += msg + System.Environment.NewLine;
            Application.DoEvents();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!processing)
                backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Process();
        }
    }
}