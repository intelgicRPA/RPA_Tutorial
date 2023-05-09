# Excel Automation Tutorial
In this tutorial we are reading the city name from the excel sheet using the ***Excel Automation*** command. We are fetching the 7 days weather forecast from [Visual Crossing Weather API](https://www.visualcrossing.com/weather-api) then writing it down on the excel sheet.

## Overall Steps

 1. Create excel instance
 2. Open the excel sheet and assign it with previously created excel instance
 3. Get the city name using the *Get Cell* command
 4. Add the table headers on the excel sheet using the *Set Cell* command
 5. Call the Weather API using the *Rest API Integration* command
 6. Parse the response JSON using the *Read JSON Model* and *Read JSON Array* command
 7. Loop the JSON Array using *Loop List* command and write down the date and temperature.
 8. Determine the text to write as note based on the maximum temperature using the *If Statements* command.
 9. Write that note using the *Set Cell* command.
 10. Close and save the workbook using *Close Excel Application* command
