from __main__ import app
from __main__ import config

import pandas as pd
from flask import request
import json

worksheet = pd.read_excel(
    config.CONFIG['excelPath'],
    engine='openpyxl',
)

def isRowFiltered(row, filters):

    ret = False

    for custome_filter in filters['filters']:
        if custome_filter['label'] in worksheet:
            cell = str(row[custome_filter['label']])
            for word_filter in custome_filter['filter']:
                if cell == word_filter or \
                   word_filter.startswith('*') and \
                   cell.endswith(word_filter[1:]) or \
                   word_filter.endswith('*') and \
                   cell.startswith(word_filter[:-1]):

                    if custome_filter['type'] == 1:
                        ret = True
                    else:
                        ret = False
                        return ret

                else:
                    if custome_filter['type'] == 1:
                        ret = False
                        return ret
                    else:
                        ret = True

    return ret

def getRowWithColumn(row, filters):
    ret = {}
    for custome_filter in filters['filters']:
        if custome_filter['label'] in worksheet:
            ret[custome_filter['label']] = str(row[custome_filter['label']])
    return ret

@app.route('/excel', methods=['GET'])
def getExcel():
    try:
        filters = ''

        with open('./config/filter.json') as filter_file:
            filters = filter_file.read()

        filters = json.loads(filters)

        ret = {
            'header': [],
            'data': []
        }

        for custome_filter in filters['filters']:
            if custome_filter['label'] in worksheet:
                ret['header'].append(custome_filter['label'])

        for row in worksheet.iterrows():
            if isRowFiltered(row[1], filters):
                ret['data'].append(getRowWithColumn(row[1], filters))

        return ret

    except Exception as err:
        return err, 5001

@app.route('/excel/header', methods=['GET'])
def getAllHeader():
    try:

        ret = []
        for header in worksheet:
            ret.append(header)
        return {'headers': ret}

    except Exception as err:
        return err, 500

@app.route('/excel/filters', methods=['GET'])
def getFilters():
    try:
        with open('./config/filter.json') as filter_file:
            return filter_file.read()
    except Exception as err:
        return err, 500

@app.route('/excel/filters', methods=['POST'])
def postFilters():
    try:
        print(request.json)
        newfilters = request.json
        print(newfilters)

        if not newfilters:
            return {"error": "Missing filter parameter"}, 400

        newfilters = { "filters": newfilters}

        with open('./config/filter.json', 'w') as filter_file:
            json.dump(newfilters, filter_file)
            return {"ok": True}

    except Exception as err:
        return { "error": err }, 500

@app.route('/excel/cell/autocomplete', methods=['POST'])
def getCellAutocomplete():
    try:
        header = request.json['header']
        word = request.json['word']

        if not header or not word:
            return {"error": "Missing header or word parameter"}, 400

        if header in worksheet.columns:
            ret = []
            for cell in worksheet[header]:
                if str(cell).startswith(word):
                    ret.append(str(cell))
            return { "propositions": ret }

    except Exception as err:
        return { "error": err }, 500
