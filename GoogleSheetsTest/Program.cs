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

namespace GoogleSheetsTest
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "WeatherLogImporter";

        static void Main(string[] args)
        {
            UserCredential credential = CreateGoogleCredential();

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            //String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            //String range = "Class Data!A2:E";

            string spreadsheetId = System.Configuration.ConfigurationManager.AppSettings["spreadsheetId"];
            string sheetNo = System.Configuration.ConfigurationManager.AppSettings["sheetNo"];
            string sheetName = "MB" + sheetNo;
            string range = sheetName + "!A:AR";

            var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == sheetName);

            if (sheet == null)
            {
                sheet = CreateSheet(service, spreadsheetId, sheetName, ref spreadsheet);
            }

            if (sheet == null)
            {
                Console.WriteLine("Unable to get sheet {0}", sheetName);
                return;
            }

            Console.WriteLine("Using sheet {0}", sheetName);

            string inputFilename = System.Configuration.ConfigurationManager.AppSettings["InputFile"];

            Console.WriteLine("Using input file {0}", inputFilename);

            string inputData;

            using (var stream = new FileStream(inputFilename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(stream))
            {
                inputData = reader.ReadToEnd();
                reader.Close();
            }

            if (string.IsNullOrEmpty(inputData))
            {
                Console.WriteLine("File empty!");
                return;
            }

            var inputLines = inputData.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            Console.WriteLine("Loaded lines {0}", inputLines.Length);

            var rangeValues = service.Spreadsheets.Values.Get(spreadsheetId, range).Execute();
            var rangeLines = 0;

            if (rangeValues.Values != null)
                rangeLines = rangeValues.Values.Count;

            Console.WriteLine("Range lines {0}", rangeLines);

            if (rangeLines > 0)
            {
                //we could seach by value, but we know that line numbers are the same...
                var lastLine = rangeValues.Values[rangeLines - 1];
            }

            var newData = new List<IList<object>>();

            for (int i = rangeLines; i < inputLines.Length; i++)
            {
                var splitedInputLine = inputLines[i].Split(";".ToCharArray());

                  var date = DateTime.ParseExact(splitedInputLine[0], "dd.MM.yy", CultureInfo.InvariantCulture);
                  var time = DateTime.ParseExact(splitedInputLine[1], "HH:mm", CultureInfo.InvariantCulture);
                  var timeOnly = time - time.Date;
                  var values = splitedInputLine.Skip(2).Select(x => decimal.Parse(x)).Cast<object>().ToList();

                /*  var data = new List<object>();
                  data.Add(new DateTime(date.Year, date.Month, date.Day, timeOnly.Hours, timeOnly.Minutes, 0));
                  data.Add(timeOnly);
                  data.AddRange(values);*/

                //  var data = new List<object>(splitedInputLine);

                var data = new List<object>();
                data.Add(splitedInputLine[0]);
                data.Add(splitedInputLine[1]);
                data.AddRange(values);

                newData.Add(data);
                
                if (i % 100 == 0)
                {
                    Console.Write(".");
                }
            }

            Console.WriteLine("");
            Console.WriteLine("Sending...");

            ValueRange vr = new ValueRange();
            vr.Values = newData;
            //vr.Range = range;
            //vr.MajorDimension = "ROWS";

            SpreadsheetsResource.ValuesResource.AppendRequest request = service.Spreadsheets.Values.Append(vr, spreadsheetId, range);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            var response = request.Execute();


            Console.WriteLine("DONE");


        }

        private static Sheet CreateSheet(SheetsService service, string spreadsheetId, string sheetName, ref Spreadsheet spreadsheet)
        {
            Console.WriteLine("Creating sheet {0}", sheetName);
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

        private static UserCredential CreateGoogleCredential()
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
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }
    }
}
