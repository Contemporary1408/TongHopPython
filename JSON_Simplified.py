import json
#source: https://api.weatherapi.com/v1/forecast.json?key=ef1c5c286eb442bc90782533240605&q=Haiphong&days=3&aqi=no&alerts=no
json_file = open(r"D:\pysap\forecast.json")
data = json.load(json_file)
#data1 = json.load(data['forecast']['forecastday'])
print(data['forecast']['forecastday'][0]['date']) #print 1st date in forecast
print(data['current']['condition']['text']) #print text in condition

# ################################
import urllib.request, json
with urllib.request.urlopen("https://api.weatherapi.com/v1/forecast.json?key=ef1c5c286eb442bc90782533240605&q=Haiphong&days=3&aqi=no&alerts=no") as url:
    data = json.load(url)
print(data['forecast']['forecastday'][0]['date']) #print date in forecast
print(data['current']['condition']['text'])

# #############################
# or update value to any key in json then save
import json
with open(r"D:\pysap\forecast.json") as file:
    nested_data = json.load(file)
nested_data['current']['condition']['text'] = "Trời nắng"
with open(r"D:\pysap\forecast.json2", 'w') as outfile:
    json.dump(nested_data, outfile)
