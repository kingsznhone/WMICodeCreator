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

namespace WMICodeCreatorTools
{
    public partial class WMICodeCreator
    {
        //---------------------------------------------------------------------------------------
        // The TargetComputerWindow class creates the windows form used
        // to enter in the target computer information used in the WMICodeCreator.
        // The TargetComputerWindow class takes in information (name and domain)
        // about a remote computer, or the name of a list of remote computers in the same domain.
        //---------------------------------------------------------------------------------------
        [ComVisible(false)]
        private class TargetComputerWindow : System.Windows.Forms.Form
        {
            private System.Windows.Forms.Button okButton;
            private System.Windows.Forms.Label remoteIntro;
            private System.Windows.Forms.Label computerNameLabel;
            private System.Windows.Forms.TextBox remoteComputerNameBox;
            private System.Windows.Forms.TextBox remoteComputerDomainBox;
            private System.Windows.Forms.Label computerDomainLabel;
            private System.Windows.Forms.TextBox arrayRemoteComputersBox;
            private System.Windows.Forms.Label arrayRemoteInfoLabel;
            private WMICodeCreator controlWindow;

            // Required designer variable.
            private System.ComponentModel.Container components = null;

            public TargetComputerWindow()
            {
                InitializeComponent();
            }

            //-------------------------------------------------------------------------
            // Constructor for the TargetComputerWindow class. This constructor
            // creates a pointer to the parent WMICodeCreator form.
            //-------------------------------------------------------------------------
            public TargetComputerWindow(WMICodeCreator form)
            {
                this.controlWindow = form;

                InitializeComponent();
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
                this.okButton = new System.Windows.Forms.Button();
                this.remoteIntro = new System.Windows.Forms.Label();
                this.computerNameLabel = new System.Windows.Forms.Label();
                this.remoteComputerNameBox = new System.Windows.Forms.TextBox();
                this.remoteComputerDomainBox = new System.Windows.Forms.TextBox();
                this.computerDomainLabel = new System.Windows.Forms.Label();
                this.arrayRemoteComputersBox = new System.Windows.Forms.TextBox();
                this.arrayRemoteInfoLabel = new System.Windows.Forms.Label();
                this.SuspendLayout();
                //
                // okButton
                //
                this.okButton.Location = new System.Drawing.Point(104, 224);
                this.okButton.Name = "okButton";
                this.okButton.Size = new System.Drawing.Size(136, 23);
                this.okButton.TabIndex = 0;
                this.okButton.Text = "OK";
                this.okButton.Click += new System.EventHandler(this.okButton_Click);
                //
                // remoteIntro
                //
                this.remoteIntro.Location = new System.Drawing.Point(16, 24);
                this.remoteIntro.Name = "remoteIntro";
                this.remoteIntro.Size = new System.Drawing.Size(320, 72);
                this.remoteIntro.TabIndex = 1;
                this.remoteIntro.Text = "You have selected to perform a task using WMI on a remote computer. Fill in the i" +
                    "nformation below about the remote computer. This information will be used in the" +
                    " code created by the WMI Code Creator.";
                //
                // computerNameLabel
                //
                this.computerNameLabel.Location = new System.Drawing.Point(24, 104);
                this.computerNameLabel.Name = "computerNameLabel";
                this.computerNameLabel.Size = new System.Drawing.Size(300, 16);
                this.computerNameLabel.TabIndex = 2;
                this.computerNameLabel.Text = "Full Name (or IP Address) of the Remote Computer:";
                //
                // remoteComputerNameBox
                //
                this.remoteComputerNameBox.Location = new System.Drawing.Point(24, 120);
                this.remoteComputerNameBox.Name = "remoteComputerNameBox";
                this.remoteComputerNameBox.Size = new System.Drawing.Size(288, 20);
                this.remoteComputerNameBox.TabIndex = 3;
                this.remoteComputerNameBox.Text = "FullComputerName";
                this.remoteComputerNameBox.TextChanged += new System.EventHandler(this.remoteComputerNameBox_TextChanged);
                //
                // remoteComputerDomainBox
                //
                this.remoteComputerDomainBox.Location = new System.Drawing.Point(24, 168);
                this.remoteComputerDomainBox.Name = "remoteComputerDomainBox";
                this.remoteComputerDomainBox.Size = new System.Drawing.Size(288, 20);
                this.remoteComputerDomainBox.TabIndex = 5;
                this.remoteComputerDomainBox.Text = "DOMAIN";
                this.remoteComputerDomainBox.TextChanged += new System.EventHandler(this.remoteComputerDomainBox_TextChanged);
                //
                // computerDomainLabel
                //
                this.computerDomainLabel.Location = new System.Drawing.Point(24, 152);
                this.computerDomainLabel.Name = "computerDomainLabel";
                this.computerDomainLabel.Size = new System.Drawing.Size(240, 16);
                this.computerDomainLabel.TabIndex = 4;
                this.computerDomainLabel.Text = "Remote Computer Domain:";
                //
                // arrayRemoteComputersBox
                //
                this.arrayRemoteComputersBox.Location = new System.Drawing.Point(24, 128);
                this.arrayRemoteComputersBox.Multiline = true;
                this.arrayRemoteComputersBox.Name = "arrayRemoteComputersBox";
                this.arrayRemoteComputersBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
                this.arrayRemoteComputersBox.Size = new System.Drawing.Size(288, 80);
                this.arrayRemoteComputersBox.TabIndex = 6;
                this.arrayRemoteComputersBox.Text = "";
                this.arrayRemoteComputersBox.Visible = false;
                this.arrayRemoteComputersBox.TextChanged += new System.EventHandler(this.arrayRemoteComputersBox_TextChanged);
                //
                // arrayRemoteInfoLabel
                //
                this.arrayRemoteInfoLabel.Location = new System.Drawing.Point(16, 96);
                this.arrayRemoteInfoLabel.Name = "arrayRemoteInfoLabel";
                this.arrayRemoteInfoLabel.Size = new System.Drawing.Size(320, 32);
                this.arrayRemoteInfoLabel.TabIndex = 7;
                this.arrayRemoteInfoLabel.Visible = false;
                //
                // TargetComputerWindow
                //
                this.AllowDrop = true;
                this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
                this.ClientSize = new System.Drawing.Size(344, 266);
                this.ControlBox = false;
                this.Controls.Add(this.arrayRemoteInfoLabel);
                this.Controls.Add(this.arrayRemoteComputersBox);
                this.Controls.Add(this.remoteComputerDomainBox);
                this.Controls.Add(this.computerDomainLabel);
                this.Controls.Add(this.remoteComputerNameBox);
                this.Controls.Add(this.computerNameLabel);
                this.Controls.Add(this.remoteIntro);
                this.Controls.Add(this.okButton);
                this.Name = "TargetComputerWindow";
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.Text = "Remote Computer Information";
                this.ResumeLayout(false);
            }

            //-------------------------------------------------------------------------
            // Handles the event when the user types in the name of a remote computer.
            //
            //-------------------------------------------------------------------------
            private void remoteComputerNameBox_TextChanged(object sender, System.EventArgs e)
            {
                this.controlWindow.GenerateEventCode();
                this.controlWindow.GenerateQueryCode();
                this.controlWindow.GenerateMethodCode();
            }

            //-------------------------------------------------------------------------
            // Handles the event when the user types in the domain of a remote computer.
            //
            //-------------------------------------------------------------------------
            private void remoteComputerDomainBox_TextChanged(object sender, System.EventArgs e)
            {
                this.controlWindow.GenerateEventCode();
                this.controlWindow.GenerateQueryCode();
                this.controlWindow.GenerateMethodCode();
            }

            //-------------------------------------------------------------------------
            // Handles the event when the user clicks the OK button on the form.
            //
            //-------------------------------------------------------------------------
            private void okButton_Click(object sender, System.EventArgs e)
            {
                this.Visible = false;

                this.controlWindow.GenerateEventCode();
                this.controlWindow.GenerateQueryCode();
                this.controlWindow.GenerateMethodCode();
            }

            //-------------------------------------------------------------------------
            // Handles the event when the user types in the names for a
            // group of remote computers.
            //-------------------------------------------------------------------------
            private void arrayRemoteComputersBox_TextChanged(object sender, System.EventArgs e)
            {
                this.controlWindow.GenerateEventCode();
                this.controlWindow.GenerateQueryCode();
                this.controlWindow.GenerateMethodCode();
            }

            //-------------------------------------------------------------------------
            // Sets the window up to allow the user to type in information
            // for a single remote computer.
            //-------------------------------------------------------------------------
            public void SetForRemoteComputerInfo()
            {
                this.remoteIntro.Text = "You have selected to perform a task using WMI on a remote computer. Fill in the i" +
                    "nformation below about the remote computer. This information will be used in the" +
                    " code created by the WMI Code Creator.";
                this.remoteIntro.Visible = true;
                this.computerDomainLabel.Visible = true;
                this.computerNameLabel.Visible = true;
                this.remoteComputerDomainBox.Visible = true;
                this.remoteComputerNameBox.Visible = true;
                this.arrayRemoteInfoLabel.Visible = false;
                this.arrayRemoteComputersBox.Visible = false;
            }

            //-------------------------------------------------------------------------
            // Sets the window up to allow the user to type in information
            // for a group of remote computers.
            //-------------------------------------------------------------------------
            public void SetForGroupComputerInfo()
            {
                this.remoteIntro.Text = "You have selected to perform a task using WMI on a group of remote computers. " +
                    "Your credentials (user name, password, and domain) will be used to connect to each computer. Make sure you are an Administrator on each computer.";
                this.remoteIntro.Visible = true;
                this.computerDomainLabel.Visible = false;
                this.computerNameLabel.Visible = false;
                this.remoteComputerDomainBox.Visible = false;
                this.remoteComputerNameBox.Visible = false;
                this.arrayRemoteInfoLabel.Visible = true;
                this.arrayRemoteInfoLabel.Text = "List one computer name per line with no blank lines between computer names.";
                this.arrayRemoteComputersBox.Visible = true;
            }

            //-------------------------------------------------------------------------
            // Gets the list of the group of remote computers.
            //
            //-------------------------------------------------------------------------
            public string GetArrayOfComputers()
            {
                return this.arrayRemoteComputersBox.Text;
            }

            //-------------------------------------------------------------------------
            // Gets the name for a single remote computer.
            //
            //-------------------------------------------------------------------------
            public string GetRemoteComputerName()
            {
                return this.remoteComputerNameBox.Text;
            }

            //-------------------------------------------------------------------------
            // Gets the domain for a single remote computer.
            //
            //-------------------------------------------------------------------------
            public string GetRemoteComputerDomain()
            {
                return this.remoteComputerDomainBox.Text;
            }
        }

    }
}