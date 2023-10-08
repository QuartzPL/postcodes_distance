import os
import time
import pandas as pd
from googlemaps import Client
from datetime import datetime, timedelta

# INFO
# takes sheet with 2 cols: Code1 and Code2 containging pairs of postcodes
# conectes with Google Directions API
# check the distances and time betwenn 2 given postcodes

def read(inputPath, sheetName):
    return pd.read_excel(inputPath, sheet_name=sheetName)

def write(outputPath, dataFrame):
    dataFrame.to_excel(outputPath, index=False)
    return print(f'Data saved: {outputPath}')

def connectAPI(apiKey):
    return Client(key=apiKey)

def calculateDistance(gmaps, code1, code2, country):
    directions = gmaps.directions(code1 + ', ' + country, code2 + ', ' + country)
    # print (directions)
    km = directions[0]["legs"][0]["distance"]["value"]/1000
    time = directions[0]["legs"][0]["duration"]["value"]/60/60
    time = (datetime.min + timedelta(seconds=int(time * 3600))).strftime('%H:%M')
    # print("%.2f km; %s h"%(km,time))
    return km, time

def updateDataFrame(gmaps, dataFrame, country):
    for index, row in dataFrame.iterrows():
        if index == 9999: break # safety break
        time.sleep(60) if index == 2999 else None # limit 3k per 1 min
        dataFrame.at[index, 'Distance'], dataFrame.at[index, 'Time'] = calculateDistance(gmaps, row['Code1'], row['Code2'], country)
        print(str(index + 1) + ' - ' + row['Code1'] + ' - ' + row['Code2'])
    return dataFrame

def main():
    apiKey = 'YOUR_GOOGLE_API_KEY'
    inputPath = r'C:\Users\Damian\pyproj\PostCodes\PostCodes.xlsx'
    sheetName = 'COMBINATIONS'
    outputPath = r'C:\Users\Damian\pyproj\PostCodes\CalculatedPostCodes.xlsx'
    dataFrame = read(inputPath, sheetName)
    # dataFrame.info()
    dataFrame['Distance'] = None
    dataFrame['Time'] = None
    gmaps = connectAPI(apiKey) # connect to Google Directions API (limit 3k per 1 min)
    country = 'Polska'
    outputDataFrame = updateDataFrame(gmaps, dataFrame, country)
    print(outputDataFrame.head(10))

    write(outputPath, outputDataFrame)

if __name__ == "__main__":
    os.system('cls')
    main()