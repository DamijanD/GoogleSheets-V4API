using System;
using System.ComponentModel;
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
        int restartAfterNRuns = 12;
        int mode = 0;

        private void Form1_Load(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;

            int timer = int.Parse(System.Configuration.ConfigurationManager.AppSettings["timer"]);
            if (System.Configuration.ConfigurationManager.AppSettings["restartAfterNRuns"] != null)
            {
                restartAfterNRuns = int.Parse(System.Configuration.ConfigurationManager.AppSettings["restartAfterNRuns"]);
            }
            mode = int.Parse(System.Configuration.ConfigurationManager.AppSettings["Mode"]);

            this.Text += " " + mode.ToString();

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
            else
            {
                LogProcessor_OnMessage("Send_Click - already processing");
            }
        }

        private void Process()
        {
            if (processing)
            {
                LogProcessor_OnMessage("Already processing! ");
                return;
            }

            try
            {
                processing = true;

                LogProcessor_OnMessage("Start: " + DateTime.Now);
                textBox1.Text = "";
                LogProcessor_OnMessage(label1.Text);

                if ((mode & 1) > 0)
                {
                    LogProcessor_OnMessage("Log proc...");
                    LogProcessor logProcessor = new LogProcessor();
                    logProcessor.OnMessage += LogProcessor_OnMessage;
                    logProcessor.Process();
                    logProcessor.OnMessage -= LogProcessor_OnMessage;
                }

                if ((mode & 2) > 0)
                {
                    LogProcessor_OnMessage("Air proc...");
                    AirDavisProcessor airp = new AirDavisProcessor();
                    airp.OnMessage += LogProcessor_OnMessage;
                    airp.Process();
                    airp.OnMessage -= LogProcessor_OnMessage;
                }

                runs++;

                //label1.Text += "... End: " + DateTime.Now + " (" + runs + ")";
                LogProcessor_OnMessage("End: " + DateTime.Now + " (" + runs + ")");
                notifyIcon1.Text = label1.Text;
            }
            catch (Exception exc)
            {
                LogProcessor_OnMessage("Process EXC: " + exc.Message);
            }
            finally
            {
                processing = false;
                LogProcessor_OnMessage("Finished");
            }
        }

        private void LogProcessor_OnMessage(string msg)
        {
            try
            {
                textBox1.Text += msg + System.Environment.NewLine;
            }
            catch { }
            System.IO.File.AppendAllText("WeatherToGoogleSheets.log", DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + " " + msg + System.Environment.NewLine);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            LogProcessor_OnMessage("timer1_Tick");

            if (!processing)
            {
                if (restartAfterNRuns > 0 && runs > restartAfterNRuns)
                {
                    LogProcessor_OnMessage("Restarting...");

                    Application.Restart();

                    //System.Diagnostics.Process.Start(Application.ExecutablePath);
                    //Application.Exit();
                    return;
                }
                backgroundWorker1.RunWorkerAsync();
            }
            else
                LogProcessor_OnMessage("timer1_Tick - already processing");

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            LogProcessor_OnMessage("backgroundWorker1_DoWork");

            Process();
        }
    }
}
