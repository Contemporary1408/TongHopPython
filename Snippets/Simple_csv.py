# Get specific item in csv file
import csv
with open(r"C:\Users\anh.doduc\Desktop\maintain.csv") as file:
    spare = list(csv.reader(file))
for i in range(1,10):
    try:
        a = spare[i][0]
        b = spare[i][1]    
        print(a,b)
    except:
        pass
