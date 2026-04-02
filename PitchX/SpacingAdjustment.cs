using System;
using System.Windows.Forms;

namespace PitchX
{
    public partial class SpacingAdjustment : Form
    {
        public float SpacingCmValue { get; private set; } = 0f;

        public SpacingAdjustment()
        {
            InitializeComponent();

            // Pressing Enter will trigger btnOK_Click
            this.AcceptButton = button1;
        }

        private void button1_Click(object sender, EventArgs e) 
        {
            if (float.TryParse(textBox1.Text, out float value))
            {
                SpacingCmValue = value;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("Please enter a valid number.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button3_Click(object sender, EventArgs e) 
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
