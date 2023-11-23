/*********************************************************************
*  This code is not supported under any Microsoft standard support program or service.
*  This code is provided AS IS without warranty of any kind. Microsoft further
*  disclaims all implied warranties including, without limitation, any implied warranties
*  of merchantability or of fitness for a particular purpose. The entire risk arising out
*  of the use or performance of this code and documentation remains with you.
*  In no event shall Microsoft, its authors, or anyone else involved in the creation,
*  production, or delivery of the code be liable for any damages whatsoever (including,
*  without limitation, damages for loss of business profits, business interruption,
*  loss of business information, or other pecuniary loss) arising out of the use of or
*  inability to use the code or documentation, even if Microsoft has been
*  advised of the possibility of such damages.
*
*  File:          WMICodeCreator.cs
*
*  Created:       May 2005
*  Version:       1.0
*
*  Description:   The WMI Code Creator is a WMI learning tool
*                 that creates WMI code examples in VBScript,
*                 C#, or VB .NET.  The examples either query for data
*                 from WMI classes, execute a method from a WMI class,
*                 or receive event notifications from WMI (or a WMI
*                 event provider).
*
* Dependencies:   There are two (that I'm aware of):
*                 1. You must run the WMI Code Creator on a WMI-enabled
*                    computer. Any Windows operating system that has
*                    the number 2000 or higher in its name, or XP,
*                    is a safe bet.
*                 2. You must have version 1.1 or higher of the .NET Framework
*                    installed on your computer.
*
********************************************************************/

using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace WMICodeCreatorTools
{
    //----------------------------------------------------------------------------
    // This class is the SplashScreenForm class, which creates a
    // start-up splash screen that appears while the WMICodeCreator is loading
    // WMI classes and gathering information about the WMI classes. The
    // splash screen contains a status bar and text.
    //----------------------------------------------------------------------------
    [ComVisible(false)]
    public class SplashScreenForm : System.Windows.Forms.Form
    {
        private System.ComponentModel.IContainer components;
        private static SplashScreenForm sSForm;
        private static Thread splashScreenThread;
        private double opacityIncrease = .05;
        private double opacityDecrease = .1;
        private System.Windows.Forms.Timer timer1;
        private const int TIMER_INTERVAL = 50;
        private System.Windows.Forms.Label statusLabel;
        private static System.Windows.Forms.ProgressBar progressBar1;
        private string introText;

        //-------------------------------------------------------------------------
        // Default constructor.
        //
        //-------------------------------------------------------------------------
        public SplashScreenForm()
        {
            //
            // Required for Windows Form Designer support.
            //
            sSForm = null;
            this.StartPosition = FormStartPosition.CenterScreen;
            splashScreenThread = null;
            InitializeComponent();

            this.Opacity = .5;
            timer1.Interval = TIMER_INTERVAL;
            timer1.Start();
            introText = "Initializing the WMI Code Creator. Loading WMI classes...";
            progressBar1.Maximum = 41;
            this.ShowInTaskbar = false;
        }

        // Clean up any resources being used.
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        // Required method for Designer support - do not modify
        // the contents of this method with the code editor.
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.statusLabel = new System.Windows.Forms.Label();
            progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            //
            // timer1
            //
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick_1);
            //
            // statusLabel
            //
            this.statusLabel.Font = new System.Drawing.Font("Courier New", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.statusLabel.Location = new System.Drawing.Point(24, 32);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(232, 72);
            this.statusLabel.TabIndex = 0;
            this.statusLabel.Text = "Initializing the WMI Code Creator. Loading WMI classes...";
            this.statusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            //
            // progressBar1
            //
            progressBar1.Location = new System.Drawing.Point(24, 115);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new System.Drawing.Size(232, 23);
            progressBar1.TabIndex = 1;
            //
            // SplashScreenForm
            //
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(292, 168);
            this.Controls.Add(progressBar1);
            this.Controls.Add(this.statusLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "SplashScreenForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "WMI Code Creator";
            this.ResumeLayout(false);
        }

        //-------------------------------------------------------------------------
        // A static entry point to launch the splash screen.
        //
        //-------------------------------------------------------------------------
        private static void ShowForm()
        {
            sSForm = new SplashScreenForm();
            Application.Run(sSForm);
        }

        //-------------------------------------------------------------------------
        // A static entry point to close the splash screen.
        //
        //-------------------------------------------------------------------------
        public static void CloseForm()
        {
            if (sSForm != null)
            {
                // Start to close.
                sSForm.opacityIncrease = -sSForm.opacityDecrease;
            }
            sSForm = null;
            splashScreenThread = null;  // Not necessary at this point.
        }

        //-------------------------------------------------------------------------
        // A static method that shows the splash screen.
        //
        //-------------------------------------------------------------------------
        public static void ShowSplashScreen()
        {
            // Only launch once.
            if (sSForm != null)
                return;
            splashScreenThread = new Thread(new ThreadStart(SplashScreenForm.ShowForm));
            splashScreenThread.IsBackground = true;
            splashScreenThread.ApartmentState = ApartmentState.STA;
            splashScreenThread.Start();
        }

        //-------------------------------------------------------------------------
        // A static method to set the status of the splash screen.
        //
        //-------------------------------------------------------------------------
        public static void SetStatus(string newStatus)
        {
            if (sSForm == null)
                return;
            sSForm.introText = newStatus;
        }

        //-------------------------------------------------------------------------
        // A static entry point to launch SplashScreen.
        //
        //-------------------------------------------------------------------------
        private void timer1_Tick_1(object sender, System.EventArgs e)
        {
            if (opacityIncrease > 0.0)
            {
                if (this.Opacity < 1)
                    this.Opacity += opacityIncrease;
            }
            else
            {
                if (this.Opacity > 0.0)
                    this.Opacity += opacityIncrease;
                else
                    this.timer1.Stop();
            }
        }

        public static void IncrementProgress()
        {
            progressBar1.Increment(1);
        }

        public static void SetProgressMax(int max)
        {
            progressBar1.Maximum = max;
        }
    }
}