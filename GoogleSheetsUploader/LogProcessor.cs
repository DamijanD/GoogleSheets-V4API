using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Linq;
using System.Globalization;

namespace GoogleSheetsUploader
{
    public class LogProcessor
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "WeatherLogImporter";

        public delegate void MessageEventHandler(string msg);
        public event MessageEventHandler OnMessage;

        public void Message(string msg)
        {
            OnMessage?.Invoke(msg);
        }

        public void Process()
        {
            UserCredential credential = CreateGoogleCredential();

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            ImportMainLog(service);

            ImportDayLog(service);

            ImportMain2Log(service);
        }



        private void ImportMainLog(SheetsService service)
        {
            Message("Main log");

            string spreadsheetId = System.Configuration.ConfigurationManager.AppSettings["spreadsheetId"];

            string inputPath = System.Configuration.ConfigurationManager.AppSettings["InputPath"];
            string filePrefix = System.Configuration.ConfigurationManager.AppSettings["FilePrefix"];

            int sheetIdFrom = int.Parse(System.Configuration.ConfigurationManager.AppSettings["From"]);
            int sheetIdTo = int.Parse(System.Configuration.ConfigurationManager.AppSettings["To"]);

            Message(string.Format("From {0} to {1}", sheetIdFrom, sheetIdTo));

            var files = System.IO.Directory.EnumerateFiles(inputPath, filePrefix + "*.txt");

            Message(string.Format("Found {0} in {1}", files.Count(), inputPath + filePrefix));

            foreach (var file in files)
            {
                Message(string.Format("Processing {0}", file));

                System.IO.FileInfo fi = new FileInfo(file);

                string sheetName = fi.Name.Replace(filePrefix, "").Replace(".txt", "");

                int sheetId = int.Parse(sheetName);
                if (sheetId < sheetIdFrom || sheetId > sheetIdTo)
                    continue;

                string range = sheetName + "!A:AR";

                var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
                var sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == sheetName);

                if (sheet == null)
                {
                    sheet = CreateSheet(service, spreadsheetId, sheetName, ref spreadsheet);
                }

                if (sheet == null)
                {
                    Message(string.Format("Unable to get sheet {0}", sheetName));
                    continue;
                }

                Message(string.Format("Using sheet {0}", sheetName));

                string inputData;

                using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    inputData = reader.ReadToEnd();
                    reader.Close();
                }

                if (string.IsNullOrEmpty(inputData))
                {
                    Message(string.Format("File empty!"));
                    continue;
                }

                var inputLines = inputData.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                Message(string.Format("Loaded lines {0}", inputLines.Length));

                var rangeValues = service.Spreadsheets.Values.Get(spreadsheetId, range).Execute();
                var rangeLines = 0;

                if (rangeValues.Values != null)
                    rangeLines = rangeValues.Values.Count;

                Message(string.Format("Range lines {0}", rangeLines));

                var newData = new List<IList<object>>();

                for (int i = rangeLines; i < inputLines.Length; i++)
                {
                    var splitedInputLine = inputLines[i].Split(";".ToCharArray());

                    var date = DateTime.ParseExact(splitedInputLine[0], "dd.MM.yy", CultureInfo.InvariantCulture);
                    var time = DateTime.ParseExact(splitedInputLine[1], "HH:mm", CultureInfo.InvariantCulture);
                    var timeOnly = time - time.Date;
                    var values = splitedInputLine.Skip(2).Select(x => decimal.Parse(x)).Cast<object>().ToList();

                    var data = new List<object>();
                    data.Add(splitedInputLine[0]);
                    data.Add(splitedInputLine[1]);
                    data.AddRange(values);

                    newData.Add(data);

                }

                if (newData.Count > 0)
                {
                    Message("Sending...");

                    ValueRange vr = new ValueRange();
                    vr.Values = newData;

                    SpreadsheetsResource.ValuesResource.AppendRequest request = service.Spreadsheets.Values.Append(vr, spreadsheetId, range);
                    request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                    request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    var response = request.Execute();

                    Message(string.Format("DONE"));
                }
                else
                {
                    Message(string.Format("No new data."));
                }
            }
        }

        private void ImportDayLog(SheetsService service)
        {
            Message("Day log");

            string spreadsheetId = System.Configuration.ConfigurationManager.AppSettings["DailySpreadsheetId"];

            string inputPath = System.Configuration.ConfigurationManager.AppSettings["DailyFile"];

            Message(string.Format("Processing day file {0}", inputPath));

            string sheetName = "dayfile";

            string range = sheetName + "!A:AT";

            var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == sheetName);

            if (sheet == null)
            {
                sheet = CreateSheet(service, spreadsheetId, sheetName, ref spreadsheet);
            }

            if (sheet == null)
            {
                Message(string.Format("Unable to get sheet {0}", sheetName));
                return;
            }

            string inputData;

            using (var stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(stream))
            {
                inputData = reader.ReadToEnd();
                reader.Close();
            }

            if (string.IsNullOrEmpty(inputData))
            {
                Message(string.Format("File empty!"));
                return;
            }

            var inputLines = inputData.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            Message(string.Format("Loaded lines {0}", inputLines.Length));

            var rangeValues = service.Spreadsheets.Values.Get(spreadsheetId, range).Execute();
            var rangeLines = 0;

            if (rangeValues.Values != null)
                rangeLines = rangeValues.Values.Count;

            Message(string.Format("Range lines {0}", rangeLines));

            var newData = new List<IList<object>>();

            for (int i = rangeLines; i < inputLines.Length; i++)
            {
                var splitedInputLine = inputLines[i].Split(";".ToCharArray());

                var date = DateTime.ParseExact(splitedInputLine[0], "dd.MM.yy", CultureInfo.InvariantCulture);
                //    var time = DateTime.ParseExact(splitedInputLine[1], "HH:mm", CultureInfo.InvariantCulture);
                //    var timeOnly = time - time.Date;
                //var values = splitedInputLine.Skip(2).Select(x => decimal.Parse(x)).Cast<object>().ToList();
                var values = splitedInputLine.Skip(1).Select(x => x).Cast<object>().ToList();

                var data = new List<object>();
                data.Add(splitedInputLine[0]);
                data.Add(decimal.Parse(splitedInputLine[1]));
                data.Add(decimal.Parse(splitedInputLine[2]));
                data.Add(splitedInputLine[3]);
                data.Add(decimal.Parse(splitedInputLine[4]));
                data.Add(splitedInputLine[5]);
                data.Add(decimal.Parse(splitedInputLine[6]));
                data.Add(splitedInputLine[7]);
                data.Add(decimal.Parse(splitedInputLine[8]));
                data.Add(splitedInputLine[9]);
                data.Add(decimal.Parse(splitedInputLine[10]));
                data.Add(splitedInputLine[11]);
                data.Add(decimal.Parse(splitedInputLine[12]));
                data.Add(splitedInputLine[13]);
                data.Add(decimal.Parse(splitedInputLine[14]));
                data.Add(decimal.Parse(splitedInputLine[15]));
                data.Add(decimal.Parse(splitedInputLine[16]));
                data.Add(decimal.Parse(splitedInputLine[17]));
                data.Add(splitedInputLine[18]);
                data.Add(decimal.Parse(splitedInputLine[19]));
                data.Add(splitedInputLine[20]);
                data.Add(decimal.Parse(splitedInputLine[21]));
                data.Add(splitedInputLine[22]);
                data.Add(decimal.Parse(splitedInputLine[23]));
                data.Add(decimal.Parse(splitedInputLine[24]));
                data.Add(decimal.Parse(splitedInputLine[25]));
                data.Add(splitedInputLine[26]);
                data.Add(decimal.Parse(splitedInputLine[27]));
                data.Add(splitedInputLine[28]);
                data.Add(decimal.Parse(splitedInputLine[29]));
                data.Add(splitedInputLine[30]);
                data.Add(decimal.Parse(splitedInputLine[31]));
                data.Add(splitedInputLine[32]);
                data.Add(decimal.Parse(splitedInputLine[33]));
                data.Add(splitedInputLine[34]);
                data.Add(decimal.Parse(splitedInputLine[35]));
                data.Add(splitedInputLine[36]);
                data.Add(decimal.Parse(splitedInputLine[37]));
                data.Add(splitedInputLine[38]);
                data.Add(decimal.Parse(splitedInputLine[39]));
                data.Add(decimal.Parse(splitedInputLine[40]));
                data.Add(decimal.Parse(splitedInputLine[41]));
                data.Add(decimal.Parse(splitedInputLine[42]));
                data.Add(splitedInputLine[43]);
                data.Add(decimal.Parse(splitedInputLine[44]));
                data.Add(splitedInputLine[45]);

                //data.Add(splitedInputLine[1]);
                //data.AddRange(values);

                newData.Add(data);

            }

            if (newData.Count > 0)
            {
                Message("Sending...");

                ValueRange vr = new ValueRange();
                vr.Values = newData;

                SpreadsheetsResource.ValuesResource.AppendRequest request = service.Spreadsheets.Values.Append(vr, spreadsheetId, range);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                var response = request.Execute();

                Message(string.Format("DONE"));
            }
            else
            {
                Message(string.Format("No new data."));
            }

        }


        private void ImportMain2Log(SheetsService service)
        {
            Message("Main2 log");

            string spreadsheetId = System.Configuration.ConfigurationManager.AppSettings["MonthSpreadsheetId"];

            string inputPath = System.Configuration.ConfigurationManager.AppSettings["InputPath"];
            string fileEnding = System.Configuration.ConfigurationManager.AppSettings["MonthFileName"];

            int sheetIdFrom = int.Parse(System.Configuration.ConfigurationManager.AppSettings["From"]);
            int sheetIdTo = int.Parse(System.Configuration.ConfigurationManager.AppSettings["To"]);

            Message(string.Format("From {0} to {1}", sheetIdFrom, sheetIdTo));

            var files = System.IO.Directory.EnumerateFiles(inputPath, "*"+ fileEnding + ".txt");

            Message(string.Format("Found {0} in {1}", files.Count(), inputPath + "*" + fileEnding + ".txt"));

            foreach (var file in files)
            {
                Message(string.Format("Processing {0}", file));

                System.IO.FileInfo fi = new FileInfo(file);

                string sheetName = fi.Name.Replace(".txt", "");

                /*int sheetId = int.Parse(sheetName);
                if (sheetId < sheetIdFrom || sheetId > sheetIdTo)
                    continue;*/

                string range = sheetName + "!A:AA";

                var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
                var sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == sheetName);

                if (sheet == null)
                {
                    sheet = CreateSheet(service, spreadsheetId, sheetName, ref spreadsheet);
                }

                if (sheet == null)
                {
                    Message(string.Format("Unable to get sheet {0}", sheetName));
                    continue;
                }

                Message(string.Format("Using sheet {0}", sheetName));

                string inputData;

                using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    inputData = reader.ReadToEnd();
                    reader.Close();
                }

                if (string.IsNullOrEmpty(inputData))
                {
                    Message(string.Format("File empty!"));
                    continue;
                }

                var inputLines = inputData.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                Message(string.Format("Loaded lines {0}", inputLines.Length));

                var rangeValues = service.Spreadsheets.Values.Get(spreadsheetId, range).Execute();
                var rangeLines = 0;

                if (rangeValues.Values != null)
                    rangeLines = rangeValues.Values.Count;

                Message(string.Format("Range lines {0}", rangeLines));

                var newData = new List<IList<object>>();

                for (int i = rangeLines; i < inputLines.Length; i++)
                {
                    var splitedInputLine = inputLines[i].Split(";".ToCharArray());

                    var date = DateTime.ParseExact(splitedInputLine[0], "dd.MM.yy", CultureInfo.InvariantCulture);
                    var time = DateTime.ParseExact(splitedInputLine[1], "HH:mm", CultureInfo.InvariantCulture);
                    var timeOnly = time - time.Date;
                    var values = splitedInputLine.Skip(2).Select(x => decimal.Parse(x)).Cast<object>().ToList();

                    var data = new List<object>();
                    data.Add(splitedInputLine[0]);
                    data.Add(splitedInputLine[1]);
                    data.AddRange(values);

                    newData.Add(data);

                }

                if (newData.Count > 0)
                {
                    Message("Sending...");

                    ValueRange vr = new ValueRange();
                    vr.Values = newData;

                    SpreadsheetsResource.ValuesResource.AppendRequest request = service.Spreadsheets.Values.Append(vr, spreadsheetId, range);
                    request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                    request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    var response = request.Execute();

                    Message(string.Format("DONE"));
                }
                else
                {
                    Message(string.Format("No new data."));
                }
            }
        }



        private Sheet CreateSheet(SheetsService service, string spreadsheetId, string sheetName, ref Spreadsheet spreadsheet)
        {
            Message(string.Format("Creating sheet {0}", sheetName));
            Request r = new Request();
            r.AddSheet = new AddSheetRequest()
            {
                Properties = new SheetProperties()
                {
                    Title = sheetName
                }
            };

            service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        AddSheet = new AddSheetRequest()
                        {
                            Properties = new SheetProperties()
                            {
                                Title = sheetName
                            }
                        }
                    }
                }
            }, spreadsheetId).Execute();

            spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
            return spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == sheetName);
        }

        private UserCredential CreateGoogleCredential()
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Message(string.Format("Credential file saved to: " + credPath));
            }

            return credential;
        }
    }
}
