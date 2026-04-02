using Newtonsoft.Json.Linq;
using PitchX;
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TLBR
{
    public class AutoUpdater
    {
        private const string GitHubOwner = "junaid042";
        private const string GitHubRepo = "PitchX";           // your repo name
        private const string MsiAssetName = "PitchX.msi";

        public static async Task CheckForUpdatesAsync()
        {
            try
            {
                var client = new System.Net.Http.HttpClient();
                client.DefaultRequestHeaders.Add("User-Agent", "PitchX-Updater");

                string url = $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";
                string json = await client.GetStringAsync(url);
                var release = JObject.Parse(json);

                string latestTag = release["tag_name"]?.ToString()?.TrimStart('v');
                string currentVersion = Assembly.GetExecutingAssembly()
                                                .GetName().Version.ToString(3);

                if (!Version.TryParse(latestTag, out var latest)) return;
                if (!Version.TryParse(currentVersion, out var current)) return;
                if (latest <= current) return;

                string notes = release["body"]?.ToString() ?? "Bug fixes and improvements.";
                var answer = MessageBox.Show(
                    $"PitchX {latestTag} is available!\n\n{notes}\n\nInstall now?",
                    "PitchX Update Available",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (answer != DialogResult.Yes) return;

                // Find MSI download URL
                string downloadUrl = null;
                foreach (var asset in (JArray)release["assets"])
                {
                    if (asset["name"]?.ToString() == MsiAssetName)
                    {
                        downloadUrl = asset["browser_download_url"]?.ToString();
                        break;
                    }
                }
                if (downloadUrl == null) return;

                // Download MSI to temp folder
                string tempMsi = Path.Combine(Path.GetTempPath(), $"PitchX_{latestTag}.msi");
                var bytes = await client.GetByteArrayAsync(downloadUrl);
                File.WriteAllBytes(tempMsi, bytes);

                // Run installer silently
                Process.Start(new ProcessStartInfo
                {
                    FileName = "msiexec.exe",
                    Arguments = $"/i \"{tempMsi}\" /passive",
                    UseShellExecute = true
                });

                // Close PowerPoint so installer can replace files
                Globals.ThisAddIn.Application.Quit();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[PitchX AutoUpdater] {ex.Message}");
                // Never crash PowerPoint — silently swallow all errors
            }
        }
    }
}