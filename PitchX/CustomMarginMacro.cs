using System;
using System.Windows.Forms;

namespace PitchX
{
    public partial class CustomMarginMacro : Form
    {
        // Public properties to retrieve values after OK
        public float MarginValueCm { get; private set; }
        public bool ApplyLeft { get; private set; }
        public bool ApplyRight { get; private set; }
        public bool ApplyTop { get; private set; }
        public bool ApplyBottom { get; private set; }

        public CustomMarginMacro()
        {
            InitializeComponent();

            // Wire up event handlers (since your Designer didn’t include them)
            button1.Click += btnCheckAll_Click;
            button2.Click += btnUncheckAll_Click;
            button3.Click += btnOK_Click;
            button4.Click += btnCancel_Click;

            // Pressing Enter will trigger btnOK_Click
            this.AcceptButton = button3;
        }

        private void CustomMarginMacro_Load(object sender, EventArgs e)
        {
            // Optional defaults
            textBox1.Text = "";
            checkBox1.Checked = true;  
            checkBox3.Checked = true; 
            checkBox4.Checked = true;  
            checkBox2.Checked = true;  

            // Optional: lock the dialog size
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnCheckAll_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = true;
            checkBox3.Checked = true; 
            checkBox4.Checked = true; 
            checkBox2.Checked = true;
        }

        private void btnUncheckAll_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox2.Checked = false;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // Validate margin value
            if (!float.TryParse(textBox1.Text, out float margin) || margin < 0)
            {
                MessageBox.Show("Please enter a valid positive numeric margin value (in cm).", "Invalid Input");
                return;
            }

            // Assign results
            MarginValueCm = margin;
            ApplyLeft = checkBox1.Checked;
            ApplyRight = checkBox3.Checked;
            ApplyTop = checkBox4.Checked;
            ApplyBottom = checkBox2.Checked;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
