using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace GoogleSheetsUploader
{
    internal class GoogleSheetUtils
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "WeatherLogImporter";

        public static UserCredential CreateGoogleCredential()
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
                //Message(string.Format("Credential file saved to: " + credPath));
            }

            return credential;
        }

        public static Sheet CreateSheet(SheetsService service, string spreadsheetId, string sheetName, ref Spreadsheet spreadsheet)
        {
            //Message(string.Format("Creating sheet {0}", sheetName));
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
    }
}
