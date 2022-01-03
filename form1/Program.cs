using System;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.Threading;

namespace form1
{
    class Program
    {
        private static readonly string[] Scopes = {SheetsService.Scope.Spreadsheets};
        private static readonly string ApplicationName = "getdata";
        private static readonly string SpreadsheetID = "17AcGes-mBiFlEBfSt5PZqr_ttT7wFSzYaz5ADoInRII";
        private static readonly string Sheet = "events";
        private static SheetsService service;

        static void Main(string[] args)
        {
            GoogleCredential credential;
            using (var streams = new FileStream("client_data.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(streams).CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            CreateEntry();
            ReadEntries();
        }

        private static void ReadEntries()
        {
            var range = $"{Sheet}!A1:C";
            var request = service.Spreadsheets.Values.Get(SpreadsheetID, range);

            var response = request.Execute();
            var values = response.Values;


            if (values != null && values.Count > 0)
            {
                Console.WriteLine($"Количество записей в таблице {values.Count - 1}");
                foreach (var row in values)
                    Console.WriteLine("{0} | {1} | {2}", row[0], row[1], row[2]);
            }
            else
            {
                Console.WriteLine("No data!");
            }
        }

        static void CreateEntry()
        {
            var range = $"{Sheet}!A:C";
            var valueRange = new ValueRange();
            var objectList = new List<object>() {3, "3:20", "4:30"};
            valueRange.Values = new List<IList<object>> {objectList};

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);
            appendRequest.ValueInputOption =
                SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }

        static void UpdateEntry()
        {
            var range = $"{Sheet}!A5";
            var valueRange = new ValueRange();
            var objectList = new List<object>() {"Исправлено"};
            valueRange.Values = new List<IList<object>> {objectList};
            var appendRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetID, range);
            appendRequest.ValueInputOption =
                SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }

        static void DeleteEntry()
        {
            var range = $"{Sheet}!A3:C";
            var requestBody = new ClearValuesRequest();
            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetID, range);
            var deleteResponse = deleteRequest.Execute();
        }
    }
}