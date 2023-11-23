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

using System;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WMICodeCreatorTools
{
    public partial class WMICodeCreator
    {
        //------------------------------------------------------------------------------
        // The InParameterWindow class is a windows form that is used by
        // the user to enter values for method in-parameters in the WMICodeCreator form.
        // An Array of InParameterWindow objects is created, with one object for
        // each method in-parameter.
        //------------------------------------------------------------------------------
        [ComVisible(false)]
        private class InParameterWindow : System.Windows.Forms.Form
        {
            private System.Windows.Forms.TextBox textBox1;
            private System.Windows.Forms.Label InputMessage;
            private System.Windows.Forms.Button OKButton;
            private string StoredValue;
            private string ParameterName;
            private bool OkButtonClicked;
            private System.Windows.Forms.Button CloseButton;
            private WMICodeCreator ParentWMIToolForm;

            // Required designer variable.
            private System.ComponentModel.Container components = null;

            //-------------------------------------------------------------------------
            // Initializes the InParameterWindow object.
            // Do not use this default constructor.
            //-------------------------------------------------------------------------
            public InParameterWindow()
            {
                //
                // Required for Windows Form Designer support.
                //
                InitializeComponent();
            }

            //-------------------------------------------------------------------------
            // Initializes the InParameterWindow object, creating a pointer
            // back to the parent WMICodeCreator form.
            //-------------------------------------------------------------------------
            public InParameterWindow(WMICodeCreator parent)
            {
                InitializeComponent();
                this.ParameterName = "";
                this.StoredValue = "";
                this.OkButtonClicked = false;
                this.ParentWMIToolForm = parent;
            }

            //-------------------------------------------------------------------------
            // Clean up any resources being used.
            //-------------------------------------------------------------------------
            protected override void Dispose(bool disposing)
            {
                for (int j = 0; j < this.ParentWMIToolForm.InParameterBox.Items.Count; j++)
                {
                    if (this.Equals(
                        this.ParentWMIToolForm.InParameterArray[j]))
                    {
                        // This deselects the in-parameter on the list and then makes a new entry into the array to
                        // replace the old in-parameter that is being deleted.
                        this.ParentWMIToolForm.InParameterBox.SetSelected(j, false);
                    }
                }

                if (disposing)
                {
                    if (components != null)
                    {
                        components.Dispose();
                    }
                }
                base.Dispose(disposing);
            }

            //-------------------------------------------------------------------------
            // Required method for Designer support - do not modify
            // the contents of this method with the code editor.
            //-------------------------------------------------------------------------
            private void InitializeComponent()
            {
                this.textBox1 = new System.Windows.Forms.TextBox();
                this.InputMessage = new System.Windows.Forms.Label();
                this.OKButton = new System.Windows.Forms.Button();
                this.CloseButton = new System.Windows.Forms.Button();
                this.SuspendLayout();
                this.TopMost = true;
                //
                // textBox1
                //
                this.textBox1.Location = new System.Drawing.Point(32, 64);
                this.textBox1.Name = "textBox1";
                this.textBox1.Size = new System.Drawing.Size(224, 20);
                this.textBox1.TabIndex = 0;
                this.textBox1.Text = "";
                this.textBox1.TextChanged += new EventHandler(TextBox_TextChanged);
                //
                // InputMessage
                //
                this.InputMessage.Location = new System.Drawing.Point(32, 16);
                this.InputMessage.Name = "InputMessage";
                this.InputMessage.Size = new System.Drawing.Size(224, 40);
                this.InputMessage.TabIndex = 1;
                this.InputMessage.Text = "";

                //
                // OKButton
                //
                this.OKButton.Location = new System.Drawing.Point(40, 104);
                this.OKButton.Name = "OKButton";
                this.OKButton.Size = new System.Drawing.Size(96, 23);
                this.OKButton.TabIndex = 2;
                this.OKButton.Text = "OK";
                this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
                //
                // CloseButton
                //
                this.CloseButton.Location = new System.Drawing.Point(152, 104);
                this.CloseButton.Name = "CloseButton";
                this.CloseButton.Size = new System.Drawing.Size(96, 23);
                this.CloseButton.TabIndex = 3;
                this.CloseButton.Text = "Cancel";
                this.CloseButton.Click += new System.EventHandler(this.CancelButton_Click);
                //
                // InParameterWindow
                //
                this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
                this.ClientSize = new System.Drawing.Size(292, 146);
                this.ControlBox = false;
                this.Controls.Add(this.CloseButton);
                this.Controls.Add(this.OKButton);
                this.Controls.Add(this.InputMessage);
                this.Controls.Add(this.textBox1);
                this.Name = "InParameterWindow";
                this.Text = "Enter in-parameter";
                this.ResumeLayout(false);
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            }

            //-------------------------------------------------------------------------
            // Handles the event when the user clicks the OK button on the
            // InParameterWindow form.
            //-------------------------------------------------------------------------
            private void OKButton_Click(object sender, System.EventArgs e)
            {
                if (this.GetParameterType().Equals("String"))
                {
                    this.StoredValue = "\"" + this.textBox1.Text + "\"";
                }
                else
                {
                    this.StoredValue = this.textBox1.Text;
                }

                this.Visible = false;
                this.OkButtonClicked = true;

                for (int j = 0; j < this.ParentWMIToolForm.InParameterBox.Items.Count; j++)
                {
                    if (this.ParameterName.Equals(
                        this.ParentWMIToolForm.InParameterBox.Items[j].ToString().Split(" ".ToCharArray())[0]))
                    {
                        string conditionName = this.ParentWMIToolForm.InParameterBox.Items[j].ToString().Split(" ".ToCharArray())[0];
                        // Updates the PropertyList_event item with the input value.
                        this.ParentWMIToolForm.InParameterBox.Items.RemoveAt(j);
                        this.ParentWMIToolForm.InParameterBox.Items.Add(conditionName + " = " + this.StoredValue);
                        this.ParentWMIToolForm.InParameterBox.Sorted = true;
                        this.ParentWMIToolForm.InParameterBox.SetSelected(j, true);
                    }
                }

                this.ParentWMIToolForm.GenerateMethodCode();
            }

            //-------------------------------------------------------------------------
            // Returns the value of OkButtonClicked
            //-------------------------------------------------------------------------
            public bool GetOkClicked()
            {
                return this.OkButtonClicked;
            }

            //-------------------------------------------------------------------------
            // Sets the value of OkButtonClicked
            //-------------------------------------------------------------------------
            public void SetOkClicked(bool setValue)
            {
                this.OkButtonClicked = setValue;
            }

            //-------------------------------------------------------------------------
            // Returns the type of the method in-parameter.
            //
            //-------------------------------------------------------------------------
            public string GetParameterType()
            {
                string type = " ";

                try
                {
                    ManagementClass c = new ManagementClass(this.ParentWMIToolForm.NamespaceValue_m.Text, this.ParentWMIToolForm.ClassList_m.Text, null);

                    ManagementBaseObject m = c.Methods[this.ParentWMIToolForm.MethodList.Text].InParameters;
                    type = m.Properties[this.ParameterName].Type.ToString();
                }
                catch (ManagementException mErr)
                {
                    if (mErr.Message.Equals("Not found "))
                        MessageBox.Show("WMI class or method not found.");
                    else
                        MessageBox.Show(mErr.Message.ToString());
                }

                return type;
            }

            //-------------------------------------------------------------------------
            // Handles the event when the user clicks the Cancel button on the
            // InParameterWindow form.
            //-------------------------------------------------------------------------
            private void CancelButton_Click(object sender, System.EventArgs e)
            {
                this.StoredValue = "";
                this.textBox1.Text = "";
                this.Visible = false;
                this.OkButtonClicked = false;

                for (int j = 0; j < this.ParentWMIToolForm.InParameterBox.Items.Count; j++)
                {
                    if (this.ParameterName.Equals(
                        this.ParentWMIToolForm.InParameterBox.Items[j].ToString().Split(" ".ToCharArray())[0]))
                    {
                        // Change the name back to no value.
                        string conditionName = this.ParentWMIToolForm.InParameterBox.Items[j].ToString().Split(" ".ToCharArray())[0];
                        // Update the PropertyList_event item with the input value.
                        this.ParentWMIToolForm.InParameterBox.Items.RemoveAt(j);
                        this.ParentWMIToolForm.InParameterBox.Items.Add(conditionName);
                        this.ParentWMIToolForm.InParameterBox.Sorted = true;

                        this.ParentWMIToolForm.InParameterBox.SetSelected(j, false);
                    }
                }

                this.ParentWMIToolForm.GenerateMethodCode();
            }

            //-------------------------------------------------------------------------
            // Handles the event when the user enters in a value for a method
            // in-parameter.
            //-------------------------------------------------------------------------
            private void TextBox_TextChanged(object sender, System.EventArgs e)
            {
                this.StoredValue = this.textBox1.Text;
                this.ParentWMIToolForm.GenerateMethodCode();
            }

            //-------------------------------------------------------------------------
            // Changes the introductory text on the
            // InParameterWindow form.
            //-------------------------------------------------------------------------
            public void ChangeText(string newText)
            {
                this.InputMessage.Text = newText;
            }

            //-------------------------------------------------------------------------
            // Returns the in-parameter value that has been entered by a user.
            //
            //-------------------------------------------------------------------------
            public string ReturnParameterValue()
            {
                return StoredValue;
            }

            //-------------------------------------------------------------------------
            // Gets the name of the method in-parameter.
            //
            //-------------------------------------------------------------------------
            public string GetParameterName()
            {
                return ParameterName;
            }

            //-------------------------------------------------------------------------
            // Sets the name of the method in-parameter.
            //
            //-------------------------------------------------------------------------
            public void SetParameterName(string inputName)
            {
                this.ParameterName = inputName;
            }
        }
    }
}