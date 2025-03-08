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
using System.IO;
using System.Xml.Serialization;
using System.Globalization;

namespace GoogleSheetsUploader
{
    internal class ArsoWaterFlowProcessor
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

                    ImportWater(service);
                 }
             }
             catch (Exception exc)
             {
                 Message("ArsoWaterFlowProcessor.Process EXC: " + exc.Message);
             }
        }

        protected T FromXml<T>(String xml)
        {
            T returnedXmlClass = default(T);

            try
            {
                using (TextReader reader = new StringReader(xml))
                {
                    try
                    {
                        returnedXmlClass =
                            (T)new XmlSerializer(typeof(T)).Deserialize(reader);
                    }
                    catch (InvalidOperationException)
                    {
                        // String passed is not XML, simply return defaultXmlClass
                    }
                }
            }
            catch (Exception ex)
            {
            }

            return returnedXmlClass;
        }

        private arsopodatki GetWaterData()
        {
           /* string ww = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<arsopodatki verzija=""1.5"">
<vir>Agencija RS za okolje</vir>
<predlagan_zajem>5 minut čez polno uro ali pol ure</predlagan_zajem>
<predlagan_zajem_perioda>30 min</predlagan_zajem_perioda>
<datum_priprave>2025-01-22 16:00</datum_priprave>
<postaja sifra=""3370"" wgs84_dolzina=""14.0999"" wgs84_sirina=""46.35697"" kota_0=""468.09"">
<reka>Natega</reka>
<merilno_mesto>Mlino</merilno_mesto>
<ime_kratko>Natega - Mlino</ime_kratko>
<datum>2025-01-22 16:00</datum>
<datum_cet>2025-01-22 16:00</datum_cet>
<vodostaj>16</vodostaj>
<prvi_vv_vodostaj/>
<drugi_vv_vodostaj/>
<tretji_vv_vodostaj/>
<vodostaj_znacilni/>
</postaja>
<postaja sifra=""3320"" wgs84_dolzina=""13.95005"" wgs84_sirina=""46.27359"" kota_0=""504.45"">
<reka>Bistrica</reka>
<merilno_mesto>Bohinjska Bistrica</merilno_mesto>
<ime_kratko>Bistrica - Bohinjska Bistrica</ime_kratko>
<datum>2025-01-22 16:00</datum>
<datum_cet>2025-01-22 16:00</datum_cet>
<vodostaj>46</vodostaj>
<pretok>1</pretok>
<prvi_vv_pretok>80.0</prvi_vv_pretok>
<drugi_vv_pretok>117.0</drugi_vv_pretok>
<tretji_vv_pretok>132.0</tretji_vv_pretok>
<pretok_znacilni>nizki pretok</pretok_znacilni>
<temp_vode>5.4</temp_vode>
</postaja>
<postaja sifra=""9420"" wgs84_dolzina=""13.53522"" wgs84_sirina=""45.60157"">
<reka>Jadransko morje</reka>
<merilno_mesto>Tržaški zaliv (Zarja)</merilno_mesto>
<ime_kratko>Jadransko morje - Tržaški zaliv</ime_kratko>
<datum>2025-01-22 16:00</datum>
<datum_cet>2025-01-22 16:00</datum_cet>
<temp_vode>9.1</temp_vode>
</postaja>
</arsopodatki>";


            return FromXml<arsopodatki>(ww);*/


            string waterUrl = System.Configuration.ConfigurationManager.AppSettings["WaterDataUrl"];

            HttpClient client = new HttpClient();

            var waterData = client.GetStringAsync(waterUrl);

            waterData.Wait();

            return FromXml<arsopodatki>(waterData.Result);
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

        private void ImportWater(SheetsService service)
        {
            Message("Water");

            string spreadsheetId = System.Configuration.ConfigurationManager.AppSettings["WaterSpreadsheetId"];

            var waterData = GetWaterData();

            if (waterData == null)
            {
                Message("Water no data received");
                return;
            }

            string sheetName = $"Water{DateTime.Now.Year}-{DateTime.Now.Month:00}";

            string range = sheetName + "!A:J";

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

            DateTime lastReceived = DateTime.MinValue;

            if (rangeValues.Values == null)
            {
                var header = new List<object>();
                header.Add("Date");
                header.Add("Time");
                header.Add("Code");
                header.Add("River");
                header.Add("Location");
                header.Add("Name");
                header.Add("Water level");
                header.Add("Flow");
                header.Add("Flow desc");
                header.Add("Temp");

                newData.Add(header);
            }
            else
            {
                var row = rangeValues.Values.LastOrDefault();
                if (row != null)
                {
                    lastReceived = DateTime.Parse(row[0] + " " + row[1]);
                }
            }

            /*postaja sifra="4222" ge_dolzina="14.111466" ge_sirina="46.043916" kota_0="474.77">
<reka>Poljanska Sora</reka>
<merilno_mesto>Žiri</merilno_mesto>
<ime_kratko>Poljanska Sora - Žiri</ime_kratko>
<datum>2022-10-09 16:00</datum>
<vodostaj>79</vodostaj>
<pretok>1.144</pretok>
<pretok_znacilni>srednji pretok</pretok_znacilni>
<temp_vode>11.5</temp_vode>
<prvi_vv_pretok>98</prvi_vv_pretok>
<drugi_vv_pretok>130</drugi_vv_pretok>
<tretji_vv_pretok>162</tretji_vv_pretok>
</postaja>*/

            var received = DateTime.Parse(waterData.datum_priprave);

            if (received > lastReceived)
            {
                var waterDataDataRecord = waterData.postaja.FirstOrDefault(x => x.ime_kratko.ToLower() == "poljanska sora - žiri");
                if (waterDataDataRecord == null)
                {
                    waterDataDataRecord = waterData.postaja.FirstOrDefault(x => x.ime_kratko.ToLower() == "poljanska sora - žiri iii");
                }

                if (waterDataDataRecord == null)
                {
                    Message("Water no data for Poljanska Sora - Žiri");
                    return;
                }

                var data = new List<object>();
                data.Add(received.ToString("dd.MM.yyyy"));
                data.Add(received.ToString("HH:mm"));
                data.Add(waterDataDataRecord.sifra);
                data.Add(waterDataDataRecord.reka);
                data.Add(waterDataDataRecord.merilno_mesto);
                data.Add(waterDataDataRecord.ime_kratko);
                //data.Add(decimal.Parse(waterDataDataRecord.vodostaj, CultureInfo.InvariantCulture));
                SafeDecimalParse(waterDataDataRecord.vodostaj, data);
                //data.Add(decimal.Parse(waterDataDataRecord.pretok, CultureInfo.InvariantCulture));
                SafeDecimalParse(waterDataDataRecord.pretok, data);
                data.Add(waterDataDataRecord.pretok_znacilni);
                //data.Add(decimal.Parse(waterDataDataRecord.temp_vode, CultureInfo.InvariantCulture));
                SafeDecimalParse(waterDataDataRecord.temp_vode, data);

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

        private static void SafeDecimalParse(string datastr, List<object> data)
        {
            if (decimal.TryParse(datastr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal decval))
                data.Add(decval);
            else
                data.Add(null);
        }
    }



    // NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
    public partial class arsopodatki
    {

        private string virField;

        private string predlagan_zajemField;

        private string predlagan_zajem_periodaField;

        private string datum_pripraveField;

        private arsopodatkiPostaja[] postajaField;

        private decimal verzijaField;

        /// <remarks/>
        public string vir
        {
            get
            {
                return this.virField;
            }
            set
            {
                this.virField = value;
            }
        }

        /// <remarks/>
        public string predlagan_zajem
        {
            get
            {
                return this.predlagan_zajemField;
            }
            set
            {
                this.predlagan_zajemField = value;
            }
        }

        /// <remarks/>
        public string predlagan_zajem_perioda
        {
            get
            {
                return this.predlagan_zajem_periodaField;
            }
            set
            {
                this.predlagan_zajem_periodaField = value;
            }
        }

        /// <remarks/>
        public string datum_priprave
        {
            get
            {
                return this.datum_pripraveField;
            }
            set
            {
                this.datum_pripraveField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("postaja")]
        public arsopodatkiPostaja[] postaja
        {
            get
            {
                return this.postajaField;
            }
            set
            {
                this.postajaField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public decimal verzija
        {
            get
            {
                return this.verzijaField;
            }
            set
            {
                this.verzijaField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class arsopodatkiPostaja
    {

        private string rekaField;

        private string merilno_mestoField;

        private string ime_kratkoField;

        private string datumField;

        private string vodostajField;

        private string vodostaj_znacilniField;

        private string pretokField;

        private string pretok_znacilniField;

        private string temp_vodeField;

        private string znacilna_visina_valovField;

        private string smer_valovanjaField;

        private string prvi_vv_vodostajField;

        private bool prvi_vv_vodostajFieldSpecified;

        private string drugi_vv_vodostajField;

        private bool drugi_vv_vodostajFieldSpecified;

        private string tretji_vv_vodostajField;

        private bool tretji_vv_vodostajFieldSpecified;

        private string prvi_vv_pretokField;

        private bool prvi_vv_pretokFieldSpecified;

        private string drugi_vv_pretokField;

        private bool drugi_vv_pretokFieldSpecified;

        private string tretji_vv_pretokField;

        private bool tretji_vv_pretokFieldSpecified;

        private ushort sifraField;

        private decimal ge_dolzinaField;

        private decimal ge_sirinaField;

        private decimal kota_0Field;

        private bool kota_0FieldSpecified;

        /// <remarks/>
        public string reka
        {
            get
            {
                return this.rekaField;
            }
            set
            {
                this.rekaField = value;
            }
        }

        /// <remarks/>
        public string merilno_mesto
        {
            get
            {
                return this.merilno_mestoField;
            }
            set
            {
                this.merilno_mestoField = value;
            }
        }

        /// <remarks/>
        public string ime_kratko
        {
            get
            {
                return this.ime_kratkoField;
            }
            set
            {
                this.ime_kratkoField = value;
            }
        }

        /// <remarks/>
        public string datum
        {
            get
            {
                return this.datumField;
            }
            set
            {
                this.datumField = value;
            }
        }

        /// <remarks/>
        public string vodostaj
        {
            get
            {
                return this.vodostajField;
            }
            set
            {
                this.vodostajField = value;
            }
        }

        /// <remarks/>
        public string vodostaj_znacilni
        {
            get
            {
                return this.vodostaj_znacilniField;
            }
            set
            {
                this.vodostaj_znacilniField = value;
            }
        }

        /// <remarks/>
        public string pretok
        {
            get
            {
                return this.pretokField;
            }
            set
            {
                this.pretokField = value;
            }
        }

        /// <remarks/>
        public string pretok_znacilni
        {
            get
            {
                return this.pretok_znacilniField;
            }
            set
            {
                this.pretok_znacilniField = value;
            }
        }

        /// <remarks/>
        public string temp_vode
        {
            get
            {
                return this.temp_vodeField;
            }
            set
            {
                this.temp_vodeField = value;
            }
        }

        /// <remarks/>
        public string znacilna_visina_valov
        {
            get
            {
                return this.znacilna_visina_valovField;
            }
            set
            {
                this.znacilna_visina_valovField = value;
            }
        }

        /// <remarks/>
        public string smer_valovanja
        {
            get
            {
                return this.smer_valovanjaField;
            }
            set
            {
                this.smer_valovanjaField = value;
            }
        }

        /// <remarks/>
        public string prvi_vv_vodostaj
        {
            get
            {
                return this.prvi_vv_vodostajField;
            }
            set
            {
                this.prvi_vv_vodostajField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool prvi_vv_vodostajSpecified
        {
            get
            {
                return this.prvi_vv_vodostajFieldSpecified;
            }
            set
            {
                this.prvi_vv_vodostajFieldSpecified = value;
            }
        }

        /// <remarks/>
        public string drugi_vv_vodostaj
        {
            get
            {
                return this.drugi_vv_vodostajField;
            }
            set
            {
                this.drugi_vv_vodostajField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool drugi_vv_vodostajSpecified
        {
            get
            {
                return this.drugi_vv_vodostajFieldSpecified;
            }
            set
            {
                this.drugi_vv_vodostajFieldSpecified = value;
            }
        }

        /// <remarks/>
        public string tretji_vv_vodostaj
        {
            get
            {
                return this.tretji_vv_vodostajField;
            }
            set
            {
                this.tretji_vv_vodostajField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool tretji_vv_vodostajSpecified
        {
            get
            {
                return this.tretji_vv_vodostajFieldSpecified;
            }
            set
            {
                this.tretji_vv_vodostajFieldSpecified = value;
            }
        }

        /// <remarks/>
        public string prvi_vv_pretok
        {
            get
            {
                return this.prvi_vv_pretokField;
            }
            set
            {
                this.prvi_vv_pretokField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool prvi_vv_pretokSpecified
        {
            get
            {
                return this.prvi_vv_pretokFieldSpecified;
            }
            set
            {
                this.prvi_vv_pretokFieldSpecified = value;
            }
        }

        /// <remarks/>
        public string drugi_vv_pretok
        {
            get
            {
                return this.drugi_vv_pretokField;
            }
            set
            {
                this.drugi_vv_pretokField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool drugi_vv_pretokSpecified
        {
            get
            {
                return this.drugi_vv_pretokFieldSpecified;
            }
            set
            {
                this.drugi_vv_pretokFieldSpecified = value;
            }
        }

        /// <remarks/>
        public string tretji_vv_pretok
        {
            get
            {
                return this.tretji_vv_pretokField;
            }
            set
            {
                this.tretji_vv_pretokField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool tretji_vv_pretokSpecified
        {
            get
            {
                return this.tretji_vv_pretokFieldSpecified;
            }
            set
            {
                this.tretji_vv_pretokFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ushort sifra
        {
            get
            {
                return this.sifraField;
            }
            set
            {
                this.sifraField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public decimal ge_dolzina
        {
            get
            {
                return this.ge_dolzinaField;
            }
            set
            {
                this.ge_dolzinaField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public decimal ge_sirina
        {
            get
            {
                return this.ge_sirinaField;
            }
            set
            {
                this.ge_sirinaField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public decimal kota_0
        {
            get
            {
                return this.kota_0Field;
            }
            set
            {
                this.kota_0Field = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool kota_0Specified
        {
            get
            {
                return this.kota_0FieldSpecified;
            }
            set
            {
                this.kota_0FieldSpecified = value;
            }
        }
    }




}
