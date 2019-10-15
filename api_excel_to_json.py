#!/usr/bin/env python3

from flask import Flask, request, jsonify
import xlrd
from collections import OrderedDict
import simplejson as json
import pandas as pd

app = Flask(__name__)

@app.route("/excel-to-json", methods=['POST'])
def excel_to_json():
    
    file_excel = request.files['excel']
    foo = file_excel.filename
    
    # Open the workbook and select the first worksheet
    wb = xlrd.open_workbook(foo)
    sh = wb.sheet_by_index(0)

    # List to hold dictionaries
    cars_list = []

    # Iterate through each row in worksheet and fetch values into dict
    for rownum in range(1, sh.nrows):
        cars = OrderedDict()
        row_values = sh.row_values(rownum)
        cars['car-id'] = int (row_values[0])
        cars['brand'] = row_values[1]
        cars['model'] = row_values[2]
        cars['miles'] = row_values[3]
        cars_list.append(cars)
        table = {}
        table['cars']= cars_list

    # Serialize the list of dicts to JSON
    j = json.dumps(table)

    # Write to file
    with open('data_api.json', 'w') as f:
        f.write(j)

    return 'Convert to Json Success'