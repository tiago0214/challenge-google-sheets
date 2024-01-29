using Google.Apis.Auth.OAuth2;
using System.IO;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using System.Reflection;
using Newtonsoft.Json;


class Program
{
    static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static readonly string AplicationName = "Chalenge";
    static readonly string SpreadSheetId = "1k9gV8uKxdwXVjkHchgntjqQDxQFm4kChLVrJezFnVso";
    static readonly string Sheet = "engenharia_de_software";

    static SheetsService service;

    static void Main(string[] args)
    {
        GoogleCredential credential;

        using (var stream = new FileStream(@"secret\client_secrets.json", FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream)
                .CreateScoped(Scopes);
        }

        service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = AplicationName
        });
        UpdateEntry();
    }
    static void UpdateEntry()
    {
        var startRow = 4; // The starting row of the table
        var rangeRead = $"{Sheet}!C{startRow}:F";
        var request = service.Spreadsheets.Values.Get(SpreadSheetId, rangeRead);

        var response = request.Execute();
        var values = response.Values;

        if (values != null && values.Count > 0)
        {
            // Iterate over the rows returned by the query
            foreach (var row in values)
            {
                var missTheClass = Convert.ToInt32(row[0]); // Valor da coluna C
                var grades = 0;

                // Iterate over cells D, E, F in the current row
                for (int i = 1; i < 4 && i < row.Count; i++) // Start from index 1 since index 0 is column C
                {
                    grades += Convert.ToInt32(row[i]);
                }

                
                var average = grades / 3;

                
                if (missTheClass > 15)
                {
                    WriteToColumnsGAndH(startRow, "Reprovado por falta", 0);
                }
                else if(average < 50)
                {
                    WriteToColumnsGAndH(startRow, "Reprovado por nota", 0);
                }
                else if (average >= 70)
                {
                    WriteToColumnsGAndH(startRow, "Aprovado", 0);
                }
                else if (average >= 50 && average <= 70)
                {
                    WriteToColumnsGAndH(startRow, "Exame final", (average + 70) / 2);
                }
                else
                {
                    WriteToColumnsGAndH(startRow, grades, 0);
                }
                startRow++;
            }
        }
    }

    // Method to write to columns G and H
    static void WriteToColumnsGAndH(int row, object valueG, object valueH)
    {
        // Write to column G
        var valueRangeG = new ValueRange();
        var objectListG = new List<object> { valueG };
        valueRangeG.Values = new List<IList<object>> { objectListG };
        var rangeEscritaG = $"{Sheet}!G{row}";

        var updateRequestG = service.Spreadsheets.Values.Update(valueRangeG, SpreadSheetId, rangeEscritaG);
        updateRequestG.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        var appendResponseG = updateRequestG.Execute();

        // Write to column H
        var valueRangeH = new ValueRange();
        var objectListH = new List<object> { valueH };
        valueRangeH.Values = new List<IList<object>> { objectListH };
        var rangeEscritaH = $"{Sheet}!H{row}";

        var updateRequestH = service.Spreadsheets.Values.Update(valueRangeH, SpreadSheetId, rangeEscritaH);
        updateRequestH.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        var appendResponseH = updateRequestH.Execute();
    }


}
