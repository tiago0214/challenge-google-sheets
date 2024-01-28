using Google.Apis.Auth.OAuth2;
using System.IO;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;

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
        using(var stream = new FileStream($@"C:\dev\GoogleSheetsAndCsharp\client_secrets.json",FileMode.Open, FileAccess.Read)) 
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
        var startRow = 4; // A linha inicial da tabela
        var rangeLeitura = $"{Sheet}!C{startRow}:F";
        var request = service.Spreadsheets.Values.Get(SpreadSheetId, rangeLeitura);

        var response = request.Execute();
        var values = response.Values;

        if (values != null && values.Count > 0)
        {
            // Itera sobre as linhas retornadas pela consulta
            foreach (var row in values)
            {
                var falta = Convert.ToInt32(row[0]); // Valor da coluna C
                var notas = 0;

                // Itera sobre as células D, E, F na linha atual
                for (int i = 1; i < 4 && i < row.Count; i++) // Começa do índice 1, pois o índice 0 é a coluna C
                {
                    notas += Convert.ToInt32(row[i]);
                }

                
                var media = notas / 3;

                
                if (falta > 15)
                {
                    WriteToColumnsGAndH(startRow, "Reprovado por falta", 0);
                }
                else if( media < 50)
                {
                    WriteToColumnsGAndH(startRow, "Reprovado por nota", 0);
                }
                else if (media > 70)
                {
                    WriteToColumnsGAndH(startRow, "Aprovado", 0);
                }
                else if (media >= 50 && media <= 70)
                {
                    WriteToColumnsGAndH(startRow, "Exame final", (media + 70) / 2);
                }
                else
                {
                    WriteToColumnsGAndH(startRow, notas, 0);
                }
                startRow++;
            }
        }
    }

    // method that write in G and H colunns
    static void WriteToColumnsGAndH(int row, object valueG, object valueH)
    {
        // write on G
        var valueRangeG = new ValueRange();
        var objectListG = new List<object> { valueG };
        valueRangeG.Values = new List<IList<object>> { objectListG };
        var rangeEscritaG = $"{Sheet}!G{row}";

        var updateRequestG = service.Spreadsheets.Values.Update(valueRangeG, SpreadSheetId, rangeEscritaG);
        updateRequestG.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        var appendResponseG = updateRequestG.Execute();

        // write on H
        var valueRangeH = new ValueRange();
        var objectListH = new List<object> { valueH };
        valueRangeH.Values = new List<IList<object>> { objectListH };
        var rangeEscritaH = $"{Sheet}!H{row}";

        var updateRequestH = service.Spreadsheets.Values.Update(valueRangeH, SpreadSheetId, rangeEscritaH);
        updateRequestH.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        var appendResponseH = updateRequestH.Execute();
    }


}
