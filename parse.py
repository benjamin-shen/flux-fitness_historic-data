#!/usr/bin/python3.7
import sys
import os
import json as js
import numpy as np
import pandas as pd
import xlrd
import xlwt
import glob
from datetime import datetime
from dateutil.parser import parse

# spreadsheet locations
SHEET_NAME = "Sheet1"
HEADER = 7
CARDIO_COLS = "B:H"
WEIGHTS_COLS = "K:Q"
START_ROW = 0
END_ROW = 35

# write destinations
CARDIO = "cardio.json"
WEIGHTS = "weights.json"


# time mapping from 6:15am to 11:45pm in 30min intervals
# I used a list because the keys are indices (ints 0 through 35)
times = [
    "6:15am",
    "6:45am",
    "7:15am",
    "7:45am",
    "8:15am",
    "8:45am",
    "9:15am",
    "9:45am",
    "10:15am",
    "10:45am",
    "11:15am",
    "11:45am",
    "12:15pm",
    "12:45pm",
    "1:15pm",
    "1:45pm",
    "2:15pm",
    "2:45pm",
    "3:15pm",
    "3:45pm",
    "4:15pm",
    "4:45pm",
    "5:15pm",
    "5:45pm",
    "6:15pm",
    "6:45pm",
    "7:15pm",
    "7:45pm",
    "8:15pm",
    "8:45pm",
    "9:15pm",
    "9:45pm",
    "10:15pm",
    "10:45pm",
    "11:15pm",
    "11:45pm",
]
# day mapping from Monday to Sunday
days = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
]

# error handling
errorMessage = "Error: Main method terminated unexpectedly.\nPlease make sure your Python version is 3.7 or above."

# create table from spreadsheet
def createTable(worksheet):
    sheet = worksheet.loc[0:35]
    table = sheet.to_json(orient="table")
    return table

# create json from data
def createJson(data,json):
    global errorMessage

    header = [k for k in data[0]][1:]
    dayMap = {}
    for i in range(7):
        dayMap[header[i]] = days[i]

    for m in data: # dictionary of dates:count
        i = m.pop("index")
        for key in m:
            try:
                d = key
                if d[-2:] == ".1":
                    d = d[:-2]
                if "." in d[:10]:
                    d = d.replace(".","/",2)
                d = str(parse(d).strftime('%Y-%m-%d'))
            except:
                errorMessage = "Error: Invalid date format."
                raise Exception
            json["history"][d] = {
                "date": key, # temp
                "day": dayMap[key],
                "data": {}
            }
    for d in json["history"]:
        try:
            date = json["history"][d].pop("date")
            for t in times:
                count = data[times.index(t)][date]
                json["history"][d]["data"][t] = count if count else -1
        except:
            pass
    return json

def tableToJson(folder):
    global errorMessage

    files = glob.glob(folder + "/*.xls")
    if len(files) == 0:
        errorMessage = "Error: Folder has no content."
        raise Exception

    cardioJson = {
        "id": "noyes",
        "history": {}
    }
    weightsJson = {
        "id": "noyes",
        "history": {}
    }
    for filename in files:
        # cardio
        cardioSheet = pd.read_excel(filename,SHEET_NAME,usecols=CARDIO_COLS,header=HEADER)
        cardioTable = createTable(cardioSheet)
        cardioDict = js.loads(cardioTable)
        cardioData = cardioDict["data"]
        if len(cardioData) != 36 or len(cardioData[0]) != 8:
            errorMessage = "Error: Invalid cardio spreadsheet format."
            raise Exception
        cardioJson = createJson(cardioData,cardioJson)

        # weights
        weightsSheet = pd.read_excel(filename,SHEET_NAME,usecols=WEIGHTS_COLS,header=HEADER)
        weightsTable = createTable(weightsSheet)
        weightsDict = js.loads(weightsTable)
        weightsData = weightsDict["data"]
        if len(weightsData) != 36 or len(weightsData[0]) != 8:
            errorMessage = "Error: Invalid weights spreadsheet format."
            raise Exception
        weightsJson = createJson(weightsData,weightsJson)

    f = open(CARDIO, "w")
    f.write(js.dumps(cardioJson, indent=2))
    f.close()

    f = open(WEIGHTS, "w")
    f.write(js.dumps(weightsJson, indent=2))
    f.close()

def main(folder):
    print(os.path.basename(__file__) + " is running...\n")
    try:
        tableToJson(folder)
        print("JSON files created.")
    except Exception:
        print(errorMessage)
    except:
        print("Error: Unknown reason.")

if __name__ == "__main__":
    try:
        folder = sys.argv[1]
        main(folder)
    except:
        print("Usage: python parse.py <folder>")

# TODO: dynamic gym id, autoname id_startDate_endDate.json
