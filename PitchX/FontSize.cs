using System;
using System.Windows.Forms;

namespace PitchX
{
    public partial class FontSize : Form
    {
        public int SelectedFontSize { get; private set; } = 0;

        public FontSize()
        {
            InitializeComponent();

            // Wire up button events
            button1.Click += button1_Click; 
            button3.Click += button3_Click;

            // Pressing Enter will trigger btnOK_Click
            this.AcceptButton = button1;
        }

        // OK button
        private void button1_Click(object sender, EventArgs e)
        {
            if (int.TryParse(textBox1.Text.Trim(), out int fontSize) && fontSize > 0)
            {
                SelectedFontSize = fontSize;
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show("Please enter a valid font size greater than 0.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // Cancel button
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
