using System;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Diagnostics;


namespace SystemTrayApp
{
    public partial class AppWindow : Form
    {

        TcpListener listener;
        TextBox resultsTextBox;

        public AppWindow()
        {
            InitializeComponent();
            this.CenterToScreen();
            this.Load += (sender, e) =>
            {
                MessageBox.Show("Now Listening for Screen Pop Requests.", "Listener Started", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            // Initializing Results textBox
            resultsTextBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ReadOnly = true,
            };
            this.Controls.Add(resultsTextBox);

            // Initializing System Tray Icon
            this.Icon = Properties.Resources.Default;
            this.SystemTrayIcon.Icon = Properties.Resources.Default;
            this.SystemTrayIcon.Text = "Screen Pop Service";
            this.SystemTrayIcon.Visible = true;

            // Starting the Listener code
            Task.Run(() => StartServer());

            // Initializing Context Menu for Right Click
            ContextMenu menu = new ContextMenu();
            menu.MenuItems.Add("Exit", ContextMenuExit);
            this.SystemTrayIcon.ContextMenu = menu;

            // Event Handlers for window
            this.Resize += WindowResize;
            this.FormClosing += WindowClosing;
            this.WindowState = FormWindowState.Minimized;
            this.Hide();

        }

        private void SystemTrayIconDoubleClick(object sender, MouseEventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void ContextMenuExit(object sender, EventArgs e)
        {
            this.SystemTrayIcon.Visible = false;
            Application.Exit();
            Environment.Exit(0);
        }

        private void WindowResize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
            }
        }

        private void WindowClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }

        private void StartServer()
        {
            try
            {
                listener = new TcpListener(IPAddress.Loopback, 4545);
                listener.Start();

                while (true)
                {
                    TcpClient client = listener.AcceptTcpClient();
                    Task.Run(() => HandleClient(client));
                }
            }
            catch (Exception ex)
            {
                WriteToFile($"An exception occurred: {ex.Message}");
            }
        }

        private void HandleClient(TcpClient client)
        {
            try
            {
                var stream = client.GetStream();
                var reader = new StreamReader(stream);
                var requestLine = reader.ReadLine();

                if (string.IsNullOrEmpty(requestLine)) return;

                var tokens = requestLine.Split(' ');

                if (tokens.Length < 2) return;

                var url = tokens[1];
                var uri = new Uri("http://localhost" + url);
                var queryString = uri.Query.TrimStart('?');

                var results = ExtractResults(queryString);
                var ani = ExtractAni(queryString);

                if (results == null || ani == null) return;

                ProcessResults(results, ani, queryString);
            }
            catch (Exception ex)
            {
                LogError($"Error occurred while handling client: {ex.Message}");
            }
            finally
            {
                client.Close();
            }
        }

        private string ExtractResults(string queryString) =>
            GetParameterFromQueryString(queryString, "results");

        private string ExtractAni(string queryString)
        {
            int startIdx = queryString.IndexOf("&") + 1;
            int endIdx = queryString.IndexOf("&", startIdx);
            endIdx = endIdx == -1 ? queryString.Length : endIdx;

            return WebUtility.UrlDecode(queryString.Substring(startIdx, endIdx - startIdx))
                .Replace("%2b", "+");
        }

        private void ProcessResults(string results, string ani, string queryString)
        {
            int numOfResults = int.Parse(results);
            this.Invoke((MethodInvoker)delegate
            {
                switch (numOfResults)
                {
                    case 0:
                        ShowNoMatchMessage(ani);
                        break;
                    case 1:
                        ProcessSingleMatch(ani, queryString);
                        break;
                    default:
                        ShowMultipleMatches(ani, numOfResults, queryString);
                        break;
                }
            });
        }

        private void ShowNoMatchMessage(string ani)
        {
            MessageBox.Show($"Incoming Call: {ani}\n\nNo contact record was found.",
                            "No Match",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        private void ProcessSingleMatch(string ani, string queryString)
        {
            var accountInfo = GetParameterFromQueryString(queryString, "r1");
            var accountInfoParts = accountInfo.Split('#');
            string name = $"{accountInfoParts[0]} {accountInfoParts[1]}";
            string accountNumber = accountInfoParts[2];

            string message = $"Call ANI: {ani}\nMatching Contacts: 1\n\nname: {name}\naccount number: {accountNumber}";

            MessageBox.Show(message,
                            "Matching Contact Information",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

            Clipboard.SetText(accountNumber);
            SendKeys.Send("^v");
        }

        private void ShowMultipleMatches(string ani, int numOfResults, string queryString)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Incoming Call ANI: {ani}");
            sb.AppendLine($"Matching Contacts: {numOfResults}");
            sb.AppendLine("");

            for (int i = 1; i <= numOfResults; i++)
            {
                string key = "r" + i;
                var accountInfo = GetParameterFromQueryString(queryString, key);
                var accountInfoParts = accountInfo.Split('#');
                string name = $"{accountInfoParts[0]} {accountInfoParts[1]}";
                string accountNumber = accountInfoParts[2];

                sb.AppendLine($"Name: {name}");
                sb.AppendLine($"Account Number: {accountNumber}");
                sb.AppendLine("");
            }

            var resultsForm = new Form
            {
                Text = "Multiple Matches Found",
                Width = 500,
                Height = 300,
            };

            var resultsTextBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Text = sb.ToString(),
                ScrollBars = ScrollBars.Vertical,
            };

            resultsForm.Controls.Add(resultsTextBox);
            resultsForm.Show(this);
        }


        // Method to extract parameter value from query string
        private string GetParameterFromQueryString(string queryString, string parameterName)
        {
            string[] parameters = queryString.TrimStart('?').Split('&');
            foreach (string parameter in parameters)
            {
                string[] keyValue = parameter.Split('=');
                if (keyValue.Length == 2 && keyValue[0] == parameterName)
                {
                    return WebUtility.UrlDecode(keyValue[1]);
                }
            }
            return null;
        }

        public void WriteToFile(string message)
        {
            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                string dateStr = DateTime.Now.Date.ToString("yyyy-MM-dd");
                string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + dateStr;

                if (!File.Exists(filePath))
                {
                    using (StreamWriter sw = File.CreateText(filePath))
                    {
                        sw.WriteLine(message);
                    }
                }
                else
                {
                    using (StreamWriter sw = File.AppendText(filePath))
                    {
                        sw.WriteLine(message);
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("MyServiceName", ex.Message, EventLogEntryType.Error);
            }
        }

        private void LogError(string message)
        {
            try
            {
                var dateStr = DateTime.Now.ToString("yyyy-MM-dd");
                var logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", $"ServiceLog_{dateStr}.log");

                Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));

                using (var writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine(message);

                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("ServiceName", ex.Message, EventLogEntryType.Error);
            }
        }

    }
}



