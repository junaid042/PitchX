using System;
using System.Windows.Forms;

namespace PitchX
{
    public partial class FormatTableForm : Form
    {
        public int HeadingRows { get; private set; } = 1;
        public int TotalRows { get; private set; } = 0;
        public string BandedOption { get; private set; } = "Rows";
        public bool FirstColumnBold { get; private set; } = false;

        public FormatTableForm()
        {
            InitializeComponent();

            // Wire up event handlers
            button1.Click += button1_Click;
            button2.Click += button2_Click;

            // Set default selections when form loads
            this.Load += FormatTableForm_Load;

            // Pressing Enter will trigger btnOK_Click
            this.AcceptButton = button1;
        }

        // Set default selected options
        private void FormatTableForm_Load(object sender, EventArgs e)
        {
            radioButton4.Checked = true;
            radioButton6.Checked = true;
            radioButton8.Checked = true;
            radioButton10.Checked = true;
        }

        // Apply button click
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // Heading Rows
                if (radioButton1.Checked) HeadingRows = 0;      
                else if (radioButton4.Checked) HeadingRows = 1; 
                else if (radioButton2.Checked) HeadingRows = 2; 
                else if (radioButton3.Checked) HeadingRows = 3;  

                // Total Rows
                if (radioButton10.Checked) TotalRows = 0;        
                else if (radioButton9.Checked) TotalRows = 1;   
                else if (radioButton11.Checked) TotalRows = 2;   

                // Banding Option
                if (radioButton7.Checked) BandedOption = "Columns";
                else if (radioButton8.Checked) BandedOption = "Rows";

                // First Column Bold (Heading Column)
                if (radioButton5.Checked) FirstColumnBold = true;
                else if (radioButton6.Checked) FirstColumnBold = false; 

                // Close with OK
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying settings:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Cancel button click
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult = DialogResult.Cancel;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error closing form:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
