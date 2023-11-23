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

using System.Management;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WMICodeCreatorTools
{
    public partial class WMICodeCreator
    {
        //--------------------------------------------------------------------------------
        // The EventQueryCondition class is a windows form that is used by
        // the user to enter values for event query conditions in the WMICodeCreator form.
        // An array of EventQueryCondition objects is created, with one object for
        // each possible event query condition.
        //--------------------------------------------------------------------------------
        [ComVisible(false)]
        private class EventQueryCondition : System.Windows.Forms.Form
        {
            private System.Windows.Forms.Label InputMessage;
            private string StoredValue;
            private string ParameterName;
            private bool OkButtonClicked;
            private System.Windows.Forms.TextBox TextBox;
            private System.Windows.Forms.Button OKbutton;
            private System.Windows.Forms.Button CloseButton;
            private System.Windows.Forms.ComboBox OperatorBox;
            private WMICodeCreator ParentWMIToolForm;

            // Required designer variable.
            private System.ComponentModel.Container components = null;

            //-------------------------------------------------------------------------
            // Initializes the EventQueryCondition object.
            // This constructor should not be used.
            //-------------------------------------------------------------------------
            private EventQueryCondition()
            {
                //
                // Required for Windows Form Designer support.
                //
                InitializeComponent();

                this.OperatorBox.Items.Add("=");
                this.OperatorBox.Items.Add("<>");
                this.OperatorBox.Items.Add(">");
                this.OperatorBox.Items.Add("<");
                this.OperatorBox.Items.Add("ISA");
            }

            //-------------------------------------------------------------------------
            // Initializes the EventQueryCondition object, to create a pointer
            // back to the parent WMICodeCreator form.
            //-------------------------------------------------------------------------
            public EventQueryCondition(WMICodeCreator parent)
            {
                InitializeComponent();
                this.ParameterName = "";
                this.StoredValue = "";
                this.OkButtonClicked = false;
                this.ParentWMIToolForm = parent;
                this.OperatorBox.Items.Add("=");
                this.OperatorBox.Items.Add("<>");
                this.OperatorBox.Items.Add(">");
                this.OperatorBox.Items.Add("<");
                this.OperatorBox.Items.Add("ISA");
            }

            // Clean up any resources being used.
            protected override void Dispose(bool disposing)
            {
                for (int j = 0; j < this.ParentWMIToolForm.PropertyList_event.Items.Count; j++)
                {
                    if (this.Equals(
                        this.ParentWMIToolForm.EventConditionArray[j]))
                    {
                        // Change the name back to no value.
                        string conditionName = this.ParentWMIToolForm.PropertyList_event.Items[j].ToString().Split(" ".ToCharArray())[0];
                        // Update the PropertyList_event item with the input value.
                        this.ParentWMIToolForm.PropertyList_event.Items.RemoveAt(j);
                        this.ParentWMIToolForm.PropertyList_event.Items.Add(conditionName);
                        this.ParentWMIToolForm.PropertyList_event.Sorted = true;

                        // This deselects the in-parameter on the list and then makes a new entry into the array to
                        // replace the old in-parameter that is being deleted.
                        this.ParentWMIToolForm.PropertyList_event.SetSelected(j, false);
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

            // Required method for Designer support - do not modify
            // the contents of this method with the code editor.
            private void InitializeComponent()
            {
                this.TextBox = new System.Windows.Forms.TextBox();
                this.InputMessage = new System.Windows.Forms.Label();
                this.OKbutton = new System.Windows.Forms.Button();
                this.CloseButton = new System.Windows.Forms.Button();
                this.OperatorBox = new System.Windows.Forms.ComboBox();
                this.SuspendLayout();
                //
                // TextBox
                //
                this.TextBox.Location = new System.Drawing.Point(112, 64);
                this.TextBox.Name = "TextBox";
                this.TextBox.Size = new System.Drawing.Size(152, 20);
                this.TextBox.TabIndex = 0;
                this.TextBox.Text = "";
                this.TextBox.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
                //
                // InputMessage
                //
                this.InputMessage.Location = new System.Drawing.Point(32, 16);
                this.InputMessage.Name = "InputMessage";
                this.InputMessage.Size = new System.Drawing.Size(240, 40);
                this.InputMessage.TabIndex = 1;
                this.InputMessage.Text = "";
                //
                // OKbutton
                //
                this.OKbutton.Location = new System.Drawing.Point(40, 104);
                this.OKbutton.Name = "OKbutton";
                this.OKbutton.Size = new System.Drawing.Size(96, 23);
                this.OKbutton.TabIndex = 2;
                this.OKbutton.Text = "OK";
                this.OKbutton.Click += new System.EventHandler(this.OKButton_Click);
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
                // OperatorBox
                //
                this.OperatorBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
                this.OperatorBox.Location = new System.Drawing.Point(32, 64);
                this.OperatorBox.Name = "OperatorBox";
                this.OperatorBox.Size = new System.Drawing.Size(56, 21);
                this.OperatorBox.TabIndex = 4;
                this.OperatorBox.Text = "=";
                //
                // EventQueryCondition
                //
                this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
                this.ClientSize = new System.Drawing.Size(296, 146);
                this.ControlBox = false;
                this.Controls.Add(this.OperatorBox);
                this.Controls.Add(this.CloseButton);
                this.Controls.Add(this.OKbutton);
                this.Controls.Add(this.InputMessage);
                this.Controls.Add(this.TextBox);
                this.Name = "EventQueryCondition";
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.Text = "Enter property value";
                this.ResumeLayout(false);
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
            // Handles the event when the user clicks the OK button on the
            // EventQueryCondition form.
            //-------------------------------------------------------------------------
            private void OKButton_Click(object sender, System.EventArgs e)
            {
                // Check to see if it is a string value.
                // If it is a string value, add single quote marks.
                if (this.GetParameterType().Equals("String"))
                {
                    this.StoredValue = "'" + this.TextBox.Text + "'";
                }
                else
                {
                    this.StoredValue = this.TextBox.Text;
                }

                this.Visible = false;
                this.OkButtonClicked = true;

                for (int j = 0; j < this.ParentWMIToolForm.PropertyList_event.Items.Count; j++)
                {
                    if (this.ParameterName.Equals(
                        this.ParentWMIToolForm.PropertyList_event.Items[j].ToString().Split(" ".ToCharArray())[0]))
                    {
                        string conditionName = this.ParentWMIToolForm.PropertyList_event.Items[j].ToString().Split(" ".ToCharArray())[0];
                        // Update the PropertyList_event item with the input value.
                        this.ParentWMIToolForm.PropertyList_event.Items.RemoveAt(j);
                        this.ParentWMIToolForm.PropertyList_event.Items.Add(conditionName + " " + this.OperatorBox.Text + " " + this.StoredValue);
                        this.ParentWMIToolForm.PropertyList_event.Sorted = true;
                        this.ParentWMIToolForm.PropertyList_event.SetSelected(j, true);
                    }
                }

                this.ParentWMIToolForm.GenerateEventCode();
            }

            //-------------------------------------------------------------------------
            // Handles the event when the user clicks the Cancel button on the
            // EventQueryCondition form.
            //-------------------------------------------------------------------------
            private void CancelButton_Click(object sender, System.EventArgs e)
            {
                this.StoredValue = "";
                this.TextBox.Text = "";
                this.Visible = false;
                this.OkButtonClicked = false;

                for (int j = 0; j < this.ParentWMIToolForm.PropertyList_event.Items.Count; j++)
                {
                    if (this.ParameterName.Equals(
                        this.ParentWMIToolForm.PropertyList_event.Items[j].ToString().Split(" ".ToCharArray())[0]))
                    {
                        // Change the name back to no value.
                        string conditionName = this.ParentWMIToolForm.PropertyList_event.Items[j].ToString().Split(" ".ToCharArray())[0];
                        // Update the PropertyList_event item with the input value.
                        this.ParentWMIToolForm.PropertyList_event.Items.RemoveAt(j);
                        this.ParentWMIToolForm.PropertyList_event.Items.Add(conditionName);
                        this.ParentWMIToolForm.PropertyList_event.Sorted = true;

                        this.ParentWMIToolForm.PropertyList_event.SetSelected(j, false);
                    }
                }

                this.ParentWMIToolForm.GenerateEventCode();
            }

            //-------------------------------------------------------------------------
            // Handles the event when the user types in a value for an
            // event query condition form.
            //-------------------------------------------------------------------------
            private void TextBox_TextChanged(object sender, System.EventArgs e)
            {
                this.StoredValue = this.TextBox.Text;
                this.ParentWMIToolForm.GenerateEventCode();
            }

            //-------------------------------------------------------------------------
            // Changes the text on the EventQueryCondition form (used as an
            // introduction on the form).
            //-------------------------------------------------------------------------
            public void ChangeText(string newText)
            {
                this.InputMessage.Text = newText;
            }

            //-------------------------------------------------------------------------
            // Changes the value of the event query condition.
            //
            //-------------------------------------------------------------------------
            public void ChangeTextBoxValue(string textValue)
            {
                this.TextBox.Text = textValue;
            }

            //-------------------------------------------------------------------------
            // Changes the operator used in the event query condition.
            //
            //-------------------------------------------------------------------------
            public void ChangeOperator(string operatorValue)
            {
                this.OperatorBox.Text = operatorValue;
                this.OperatorBox.SelectedText = operatorValue;
            }

            //-------------------------------------------------------------------------
            // Gets the name of the parameter used in the event query condition.
            //
            //-------------------------------------------------------------------------
            public string GetParameterName()
            {
                return ParameterName;
            }

            //-------------------------------------------------------------------------
            // Sets the name of the parameter in the event query condition.
            //
            //-------------------------------------------------------------------------
            public void SetParameterName(string inputName)
            {
                this.ParameterName = inputName;
            }

            //-------------------------------------------------------------------------
            // Gets the type of the parameter used in the event query condition.
            //
            //-------------------------------------------------------------------------
            public string GetParameterType()
            {
                string type = "";
                try
                {
                    ManagementClass c = new ManagementClass(this.ParentWMIToolForm.NamespaceList_event.Text, this.ParentWMIToolForm.ClassList_event.Text, null);

                    foreach (PropertyData pData in c.Properties)
                    {
                        if (pData.Name.Equals(this.ParameterName))
                        {
                            type = pData.Type.ToString();
                        }
                    }

                    if (type.Length == 0)
                    {
                        ManagementClass c2 = new ManagementClass(this.ParentWMIToolForm.NamespaceList_event.Text, this.ParentWMIToolForm.TargetClassList_event.Text, null);

                        foreach (PropertyData p in c2.Properties)
                        {
                            if (p.Name.Equals(this.ParameterName.Split(".".ToCharArray())[1]))
                            {
                                type = p.Type.ToString();
                            }
                        }
                    }
                }
                catch (ManagementException e)
                {
                    MessageBox.Show("Error getting the type of the event class. The namespace name or event class name is incorrect.");
                }

                return type;
            }
        }

    }
}