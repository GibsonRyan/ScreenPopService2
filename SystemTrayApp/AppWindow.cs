using System;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Configuration;
using System.Collections.Generic;

namespace SystemTrayApp
{
    public partial class AppWindow : Form
    {

        TcpListener listener;
        TextBox resultsTextBox;
        private int portNumber;
        private IntPtr CubshWnd;
        private List<string> processNames;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);


        public AppWindow()
        {
            InitializeComponent();
            this.CenterToScreen();
            this.Load += (sender, e) =>
            {
                // Get port number from the user
                string input = Microsoft.VisualBasic.Interaction.InputBox("Enter Port Number:", "Configuration", "4545", -1, -1);
                if (!int.TryParse(input, out portNumber) || portNumber <= 0 || portNumber >= 65536)
                {
                    MessageBox.Show("Invalid Port Number. Application will exit.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                    return;
                }

                // Get process names from the user
                string processInput = Microsoft.VisualBasic.Interaction.InputBox("Enter Process Names (comma separated):", "Configuration", "ALLIANCEONE,putty,cuemulate", -1, -1);
                processNames = processInput.Split(',').Select(p => p.Trim()).ToList();


                MessageBox.Show("Now Listening for Screen Pop Requests on port " + portNumber, "Listener Started", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Starting the Listener code
                Task.Run(() => StartServer());
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
                listener = new TcpListener(IPAddress.Loopback, portNumber);
                listener.Start();

                while (true)
                {
                    TcpClient client = listener.AcceptTcpClient();
                    Task.Run(() => HandleClient(client));
                }
            }
            catch (Exception ex)
            {
                LogMessage($"An exception occurred: {ex.Message}");
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

                if (results == "search_exception")
                {
                    string description = GetParameterFromQueryString(queryString, "description");
                    ShowSearchExceptionMessage(description);
                    return;
                }

                ProcessResults(results, ani, queryString);
            }
            catch (Exception ex)
            {
                LogMessage($"Error occurred while handling client: {ex.Message}");
            }
            finally
            {
                client.Close();
            }
        }

        private string ExtractResults(string queryString) =>
            GetParameterFromQueryString(queryString, "results");

        private string ExtractAni(string queryString) =>
            GetParameterFromQueryString(queryString, "ani")?.Replace("%2b", "+");


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
            PerformPasteAction();
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
                Text = "Multiple Matches",
                Width = 300,
                Height = 250,
                StartPosition = FormStartPosition.CenterScreen,
                Icon = SystemIcons.Information
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

        private void ShowSearchExceptionMessage(string description)
        {
            MessageBox.Show($"Error: {description}",
                            "Search Exception",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
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

        // Get the process window for to be used for the paste
        void GetHWndFromRunningProcesses()
        {
            foreach (string targetName in processNames)
            {
                foreach (Process process in Process.GetProcesses())
                {
                    if (process.ProcessName.Equals(targetName, StringComparison.OrdinalIgnoreCase))
                    {
                        CubshWnd = process.MainWindowHandle;
                        return;
                    }
                }
            }


            CubshWnd = IntPtr.Zero;
        }

        void PasteAcctIdAndEnter()
        {
            SetForegroundWindow(CubshWnd);
            SendKeys.SendWait("^{v}");
            SendKeys.SendWait("{ENTER}");
        }

        // To be called in the Single match in order to perform the paste and enter
        void PerformPasteAction()
        {
            GetHWndFromRunningProcesses();
            if (CubshWnd != IntPtr.Zero)
            {
                PasteAcctIdAndEnter();
            }
            else
            {
                MessageBox.Show("Cannot find a window to paste account ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private void LogMessage(string message)
        {
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

                // Create directory if it does not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                string dateStr = DateTime.Now.ToString("yyyy-MM-dd");
                string fileName = $"ErrorLog_{dateStr}.log";
                string filePath = Path.Combine(path, fileName);

                // Write the log message to the file
                using (StreamWriter sw = new StreamWriter(filePath, true)) // Append mode
                {
                    sw.WriteLine(message);
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("MyServiceName", ex.Message, EventLogEntryType.Error);
            }
        }

    }
}



