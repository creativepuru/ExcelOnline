using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;


namespace ExcelOnline
{
    public partial class Form1 : Form
    {
        private SheetsService sheetsService;

        public Form1()
        {
            InitializeComponent();

            // Calling the authorization method to get the authorized SheetsService object
            sheetsService = GoogleSheetsAuthentication.GetService();
        }

        private void buttonRead_Click(object sender, EventArgs e)
        {
            this.Invoke(method: new Action(async () =>
            {
                // Reading data from Spreadsheet
                var spreadsheetId = "1mJaB-cHFzEjvJDtat8mvmrHq5ABqpsW4WR8ZNnRAyHE";
                var range = comboBox1.Text + comboBox2.Text;

                var request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = await request.ExecuteAsync();

                IList<IList<object>> values = response.Values;

                if (values != null && values.Count > 0)
                {
                    textBox1.Text = values[0][0].ToString();
                }
                else
                {
                    textBox1.Text = "No Data Found!";
                }

            }));
        }

        private void buttonWrite_Click(object sender, EventArgs e)
        {
            this.Invoke(method: new Action(async () =>
            {
                var rowData = new List<object>
                {
                    textBox2.Text
                };

                // Writing data to Google Sheet
                var spreadsheetId = "1mJaB-cHFzEjvJDtat8mvmrHq5ABqpsW4WR8ZNnRAyHE";
                var range = comboBox3.Text + comboBox4.Text;
                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { rowData }
                };

                var updateRequest = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);

                updateRequest.ValueInputOption =
                    SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

                var updateResponse = await updateRequest.ExecuteAsync();
            }));
        }

        public bool isConnect = false;

        private void buttonConnect_Click(object sender, EventArgs e)
        {
            isConnect = !isConnect;
            
            if (isConnect)
            {
                serialPort1.PortName = comboBoxCOM.Text;
                serialPort1.BaudRate = int.Parse(comboBoxBaud.Text);
                serialPort1.Open();
                timer1.Enabled = true;
                timer1.Interval = 1000;

                buttonConnect.Text = "Disconnect";
            }
            else
            {
                
                serialPort1.Close();
                timer1.Stop();
                buttonConnect.Text = "Connect";
            }
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.Invoke(method: new Action(async () =>
            {
                // Reading data from Spreadsheet
                var spreadsheetId = "1mJaB-cHFzEjvJDtat8mvmrHq5ABqpsW4WR8ZNnRAyHE";
                var range = "Sheet1!B1";

                var request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = await request.ExecuteAsync();

                IList<IList<object>> values = response.Values;

                if (values != null && values.Count > 0)
                {
                    textBox3.Text = values[0][0].ToString();
                }
                else
                {
                    ;
                }

            }));

            serialPort1.Write(textBox3.Text);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            timer1.Stop();
            serialPort1.Close();
        }
    }

    public class GoogleSheetsAuthentication
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "ABC";
        static string ClientSecretFilePath;

        static GoogleSheetsAuthentication()
        {
            string currentDirectory = Directory.GetCurrentDirectory();
            ClientSecretFilePath = Path.Combine(currentDirectory, "client_secret_529940815053-9c2s5kpb0tm5lgd7lhlmramlj1eb5rhe.apps.googleusercontent.com.json");
        }

        public static SheetsService GetService()
        {
            UserCredential credential;

            using (var stream = new FileStream(ClientSecretFilePath, FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets, Scopes, "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets service using the authorized credential
            return new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
    }
}
