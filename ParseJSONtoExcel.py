import pandas as pd
import json
#source: https://api.weatherapi.com/v1/forecast.json?key=ef1c5c286eb442bc90782533240605&q=Haiphong&days=3&aqi=no&alerts=no


# Load your nested JSON data
with open(r"D:\pysap\forecast.json") as file:
    nested_data = json.load(file)
nested_data1 = nested_data['forecast']['forecastday']

# Use json_normalize to flatten the data
df = pd.json_normalize(nested_data1,meta='date')

# Convert the flattened DataFrame to an Excel file
df.to_excel(r"D:\pysap\weat.xlsx",sheet_name="Forecast" ,engine='openpyxl', index=False)
