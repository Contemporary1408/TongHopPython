import json
#source: https://api.weatherapi.com/v1/forecast.json?key=ef1c5c286eb442bc90782533240605&q=Haiphong&days=3&aqi=no&alerts=no
json_file = open(r"D:\pysap\forecast.json")
data = json.load(json_file)
#data1 = json.load(data['forecast']['forecastday'])
print(data['forecast']['forecastday'][0]['date']) #print 1st date in forecast
print(data['current']['condition']['text']) #print text in condition

# ################################
import requests as rq
import json
src = rq.get("https://api.weatherapi.com/v1/forecast.json?key=ef1c5c286eb442bc90782533240605&q=Haiphong&days=3&aqi=no&alerts=no",headers=getHeaders())
data = json.load(data.content)
print(data['forecast']['forecastday'][0]['date']) #print date in forecast
print(data['current']['condition']['text'])
