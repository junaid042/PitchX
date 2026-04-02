using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PitchX
{
    public partial class VerifyOTP : Form
    {
        private TextBox[] otpBoxes;
        private readonly string _email;

        // New constructor that accepts email from login form
        public VerifyOTP(string email)
        {
            InitializeComponent();
            _email = email;
            this.Load += VerifyOTP_Load;
        }

        private void VerifyOTP_Load(object sender, EventArgs e)
        {
            otpBoxes = new TextBox[] { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6 };

            foreach (var box in otpBoxes)
            {
                box.MaxLength = 1;
                box.TextAlign = HorizontalAlignment.Center;
                box.Font = new System.Drawing.Font("Segoe UI", 14);

                // Only digits allowed
                box.KeyPress += (s, ev) =>
                {
                    if (!char.IsControl(ev.KeyChar) && !char.IsDigit(ev.KeyChar))
                        ev.Handled = true;
                };

                // Move to next box after one digit
                box.TextChanged += (s, ev) =>
                {
                    var current = (TextBox)s;
                    if (current.Text.Length == 1)
                    {
                        int index = Array.IndexOf(otpBoxes, current);
                        if (index < otpBoxes.Length - 1)
                            otpBoxes[index + 1].Focus();
                    }
                };

                // Handle Backspace & Paste
                box.KeyDown += (s, ev) =>
                {
                    var current = (TextBox)s;
                    int index = Array.IndexOf(otpBoxes, current);

                    // Move back on Backspace
                    if (ev.KeyCode == Keys.Back && current.Text.Length == 0 && index > 0)
                        otpBoxes[index - 1].Focus();

                    // Handle paste (Ctrl+V)
                    if (ev.Control && ev.KeyCode == Keys.V)
                    {
                        string clipboardText = Clipboard.GetText().Trim();
                        if (!string.IsNullOrEmpty(clipboardText))
                        {
                            PasteOTP(clipboardText);
                            ev.Handled = true;
                        }
                    }
                };
            }

            otpBoxes[0].Focus();
        }

        private void PasteOTP(string pasted)
        {
            var digits = pasted.Where(char.IsDigit).Take(otpBoxes.Length).ToArray();

            for (int i = 0; i < digits.Length; i++)
            {
                otpBoxes[i].Text = digits[i].ToString();
            }

            if (digits.Length < otpBoxes.Length)
            {
                otpBoxes[digits.Length].Focus();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Cancel button closes and returns to login form
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            // Combine OTP boxes into one string
            string otp = string.Join("", new[] { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6 }
                .Select(b => b.Text.Trim()));

            if (otp.Length != 6)
            {
                MessageBox.Show("Please enter a valid 6-digit OTP.", "Invalid OTP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            this.Enabled = false;

            using (var loader = new Verifying())
            {
                loader.Show();
                loader.Refresh();

                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                    using (var client = new System.Net.Http.HttpClient())
                    {
                        // Use the same email from login form
                        var payload = new { email = _email, otp };
                        string jsonBody = JsonConvert.SerializeObject(payload);
                        var content = new System.Net.Http.StringContent(jsonBody, Encoding.UTF8, "application/json");

                        string apiUrl = "https://tlbr-io-backend.vercel.app/api/v1/verify/otp";
                        var response = await client.PostAsync(apiUrl, content);
                        string responseBody = await response.Content.ReadAsStringAsync();
                        JObject result = JObject.Parse(responseBody);
                        string message = result["message"]?.ToString() ?? "";

                        loader.Close();
                        this.Enabled = true;

                        if (response.IsSuccessStatusCode && message.Contains("Welcome"))
                        {
                            MessageBox.Show(message, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            this.DialogResult = DialogResult.OK;
                            this.Close();
                        }
                        else
                        {
                            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    loader.Close();
                    this.Enabled = true;
                    MessageBox.Show($"Error verifying OTP: {ex.Message}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}