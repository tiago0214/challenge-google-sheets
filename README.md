Google Sheets Data Processing Application
Overview
This C# console application interacts with Google Sheets API to read data from a specific range, perform calculations based on certain conditions, and update the spreadsheet accordingly.

Prerequisites
Before running the application, make sure you have the following:

.NET SDK installed on your machine.
Google Sheets API credentials file (client_secrets.json) placed in the application directory. You can obtain this file by setting up a Google Cloud Platform project and enabling the Google Sheets API.
How to Run
Clone this repository to your local machine:

bash
git clone https://github.com/your-username/your-repository.git
Navigate to the project directory:

bash
cd GoogleSheetsDataProcessing
Replace the SpreadSheetId and Sheet variables in the Program.cs file with your Google Sheets spreadsheet ID and sheet name.

Build and run the application:

bash
dotnet run
What It Does
This application reads data from a specified range in a Google Sheets spreadsheet. It calculates the sum of certain columns, evaluates conditions based on absence and average grades, and updates the spreadsheet with the results in columns G and H.

If the absence is greater than 15, it marks the result as "Reprovado por falta" in column G and 0 in column H.
If the average grade is greater than 70, it marks the result as "Aprovado" in column G and 0 in column H.
If the average grade is between 50 and 70, it marks the result as "Exame final" in column G and calculates the adjusted value for column H.
Otherwise, it marks the result as the sum of the grades in column G and 0 in column H.
Feel free to modify the code or adjust the conditions to suit your specific use case.
