using System;
using System.Windows.Forms;

namespace PitchX
{
    public partial class CopyLink : Form
    {
        public CopyLink(string link)
        {
            InitializeComponent();
            textBox1.Text = link;

            // Pressing Enter will trigger btnOK_Click
            this.AcceptButton = button1;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
