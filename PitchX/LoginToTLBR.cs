using System;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;

namespace PitchX
{
    public partial class LoginToTLBR : Form
    {
        // Public class to hold login data
        public class LoginData
        {
            public string Email { get; set; }
            public string Password { get; set; }
        }
        private LoginData loginData;

        public LoginData GetLoginData()
        {
            return loginData;
        }

        public LoginToTLBR()
        {
            InitializeComponent();
            loginData = new LoginData();
        }

        private void LoginToPitchX_Load(object sender, EventArgs e)
        {
            // Optional: Set focus to email textbox
            textBox1.Focus();
        }

        // Handle Cancel button click
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        // Handle Cancel button click
        private async void button2_Click(object sender, EventArgs e)
        {
            // Trim spaces at start and end of both fields
            textBox1.Text = textBox1.Text.Trim();
            textBox2.Text = textBox2.Text.Trim();

            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("Please enter your email address.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox1.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("Please enter your password.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox2.Focus();
                return;
            }

            string email = textBox1.Text;
            string password = textBox2.Text;

            // Disable the entire form so user can't click anything
            this.Enabled = false;

            using (var loader = new Logging())
            {
                loader.Show();
                loader.Refresh();

                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                    using (var client = new System.Net.Http.HttpClient())
                    {
                        var payload = new { email, password };
                        string jsonBody = JsonConvert.SerializeObject(payload);
                        var content = new System.Net.Http.StringContent(jsonBody, Encoding.UTF8, "application/json");

                        string apiUrl = "https://tlbr-io-backend.vercel.app/api/v1/login";
                        var response = await client.PostAsync(apiUrl, content);
                        string responseBody = await response.Content.ReadAsStringAsync();
                        JObject result = JObject.Parse(responseBody);
                        string message = result["message"]?.ToString() ?? "";

                        loader.Close();
                        this.Enabled = true;

                        if (response.IsSuccessStatusCode)
                        {
                            MessageBox.Show(message, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            
                            // Store email and password for later use
                            Properties.Settings.Default.Email = email;
                            Properties.Settings.Default.Password = password;
                            Properties.Settings.Default.Save();

                            if (message == "If your email is valid, an OTP has been sent to your email address. It will expire in 5 minutes.")
                            {
                                // Close the login form
                                this.Hide();

                                // Open OTP verification dialog
                                using (var otpDialog = new VerifyOTP(email))
                                {
                                    otpDialog.StartPosition = FormStartPosition.CenterScreen;
                                    if (otpDialog.ShowDialog() == DialogResult.OK)
                                    {
                                        // OTP verified successfully
                                        MessageBox.Show("You’re now logged in!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }
                                    else
                                    {
                                        // User cancelled or failed OTP, reopen login
                                        this.Show();
                                    }
                                }
                            }
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
                    MessageBox.Show($"Error calling API: {ex.Message}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Handle Cancel button click
        private void button3_Click(object sender, EventArgs e)
        {
            OpenWebPage("https://tlbr-io.vercel.app");
        }

        // Handle Cancel button click
        private void button4_Click(object sender, EventArgs e)
        {
            OpenWebPage("https://tlbr-io.vercel.app");
        }

        private void OpenWebPage(string url)
        {
            try
            {
                // For .NET Framework
                System.Diagnostics.Process.Start(url);
            }
            catch
            {
                try
                {
                    // For .NET Core / .NET 5+
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = url,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Unable to open web page: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}