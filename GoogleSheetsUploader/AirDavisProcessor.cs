using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheetsUploader
{
    internal class AirDavisProcessor
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
             try
             {
                 UserCredential credential = GoogleSheetUtils.CreateGoogleCredential();

                 // Create Google Sheets API service.
                 using (var service = new SheetsService(new BaseClientService.Initializer()
                 {
                     HttpClientInitializer = credential,
                     ApplicationName = ApplicationName,
                 }))
                 {
                    service.HttpClient.Timeout = TimeSpan.FromMinutes(1);

                    ImportAirDavis(service);
                 }
             }
             catch (Exception exc)
             {
                 Message("AirDavisProcessor.Process EXC: " + exc.Message);
             }
        }

        private List<Root> GetAirData()
        {
            string airUrl = System.Configuration.ConfigurationManager.AppSettings["AirDataUrl"];

            HttpClient client = new HttpClient();

            var airData = client.GetFromJsonAsync<List<Root>>(airUrl);

            airData.Wait();

            return airData.Result;
        }

        private decimal GetDecimalFromJson(object o1)
        {
            var o = (System.Text.Json.JsonElement)o1;
            if (o.ValueKind == System.Text.Json.JsonValueKind.Number)
            {
                return o.GetDecimal();
            }
            else
            {
                return o.GetProperty("calculatedRisk").GetDecimal();
                //var xxxx = (AqiRiskValue)x.value;
            }
        }

        private void ImportAirDavis(SheetsService service)
        {
            Message("Air");

            string spreadsheetId = System.Configuration.ConfigurationManager.AppSettings["AirSpreadsheetId"];

            var airData = GetAirData();

            if (airData == null)
            {
                Message("Air no data received");
                return;
            }

            string sheetName = $"Air{DateTime.Now.Year}";

            string range = sheetName + "!A:AR";

            var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == sheetName);

            if (sheet == null)
            {
                sheet = GoogleSheetUtils.CreateSheet(service, spreadsheetId, sheetName, ref spreadsheet);
            }

            if (sheet == null)
            {
                Message(string.Format("Unable to get sheet {0}", sheetName));
                return;
            }

            Message(string.Format("Using sheet {0}", sheetName));

            var rangeValues = service.Spreadsheets.Values.Get(spreadsheetId, range).Execute();

            var newData = new List<IList<object>>();

            if (rangeValues.Values == null)
            {
                var header = new List<object>();
                header.Add("Date");
                header.Add("Time");
                header.Add("Temp");
                header.Add("Hum");
                header.Add("PM 1");
                header.Add("PM 2.5");
                header.Add("PM 2.5 Last 1 Hr");
                header.Add("PM 10");
                header.Add("PM 10 Last 1 Hr");
                header.Add("AQI");
                header.Add("1 Hour AQI");

                newData.Add(header);
            }

            var airDataRecord = airData[0];
            var received = DateTimeOffset.FromUnixTimeMilliseconds(airDataRecord.lastReceived.Value);

            var data = new List<object>();
            data.Add(received.ToString("dd.MM.yy"));
            data.Add(received.ToString("HH:mm"));
            data.Add(GetDecimalFromJson(airDataRecord.currConditionValues.First(x => x.sensorDataName == "Temp").value));
            data.Add(GetDecimalFromJson(airDataRecord.currConditionValues.First(x => x.sensorDataName == "Hum").value));
            data.Add(GetDecimalFromJson(airDataRecord.currConditionValues.First(x => x.sensorDataName == "PM 1").value));
            data.Add(GetDecimalFromJson(airDataRecord.currConditionValues.First(x => x.sensorDataName == "PM 2.5").value));
            data.Add(GetDecimalFromJson(airDataRecord.currConditionValues.First(x => x.sensorDataName == "PM 2.5 Last 1 Hr").value));
            data.Add(GetDecimalFromJson(airDataRecord.currConditionValues.First(x => x.sensorDataName == "PM 10").value));
            data.Add(GetDecimalFromJson(airDataRecord.currConditionValues.First(x => x.sensorDataName == "PM 10 Last 1 Hr").value));
            data.Add(GetDecimalFromJson(airDataRecord.currConditionValues.First(x => x.sensorDataName == "AQI").value));
            data.Add(GetDecimalFromJson(airDataRecord.currConditionValues.First(x => x.sensorDataName == "1 Hour AQI").value));

            newData.Add(data);

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

    public class AqiRisk
    {
        public decimal fconcentrationRangeHigh { get; set; }
        public decimal fconcentrationRangeLow { get; set; }
        public decimal fseverityIndexRangeHigh { get; set; }
        public decimal fseverityIndexRangeLow { get; set; }
        public string slangOfSeverityDescription { get; set; }
        public int ischemeId { get; set; }
        public int ipmId { get; set; }
        public decimal fseverityIndexRangePseudoHigh { get; set; }
        public string scolorDefintion { get; set; }
        public int iseverityIndexRangeHigh { get; set; }
        public int iseverityIndexRangeLow { get; set; }
        public int inumberOfFractionalDigits { get; set; }
        public string scolorDescription { get; set; }
        public int iseverityLevel { get; set; }
        public string sseverityDescription { get; set; }
        public int idescriptionGroupId { get; set; }
    }

    public class AqiRiskValue
    {
        public int calculatedRisk { get; set; }
        public AqiRisk aqiRisk { get; set; }
    }
    
    public class SensorValues
    {
        public int? sensorDataTypeId { get; set; }
        public string sensorDataName { get; set; }
        public object displayName { get; set; }
        public object reportedValue { get; set; }
        public object value { get; set; }
        public string convertedValue { get; set; }
        public string category { get; set; }
        public int? assocSensorDataTypeId { get; set; }
        public int? sortOrder { get; set; }
        public string unitLabel { get; set; }
    }

    public class TimeSeriesValues
    {
    }

    public class TimeSeriesWeekValues
    {
    }

    public class AdditionalData
    {
        public int? lastUpdated { get; set; }
        public string AQ_ENVIRONMENT { get; set; }
        public string tz { get; set; }
        public int? logicalSensorId { get; set; }
        public int? sensorProductTypeId { get; set; }
    }

    public class Root
    {
        public string ownerName { get; set; }
        public long? lastReceived { get; set; }
        public List<SensorValues> currConditionValues { get; set; }
        public List<SensorValues> highLowValues { get; set; }
        //public List<object> aggregatedValues { get; set; }
        public TimeSeriesValues timeSeriesValues { get; set; }
        public TimeSeriesWeekValues timeSeriesWeekValues { get; set; }
        public AdditionalData additionalData { get; set; }
    }




}
