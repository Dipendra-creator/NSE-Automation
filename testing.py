import requests
import time
import json
import pandas as pd
import os
import webbrowser as wb
import datetime
from datetime import date
import xlsxwriter
import flask
from flask import request, jsonify


def createDir(parentDirectory, childDirectory):
    try:
        # Path
        path = os.path.join(parentDirectory, childDirectory)

        # Create the directory
        os.mkdir(path)
        print("Directory '% s' created" % childDirectory)

    except FileExistsError:
        print("Directory '% s' created" % childDirectory + " Earlier")


while True:

    # Returns the current local date
    today = date.today()
    # p_d = 'C:/PyProgram/' + str(today)
    # print(p_d)

    # Current Time
    cTime = datetime.datetime.now().strftime("%I%M")
    cTimeH = datetime.datetime.now().strftime("%H")
    print(cTimeH)

    try:
        url = 'https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY'
        headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36',
            'accept-language': 'en-US,en;q=0.9,bn;q=0.8', 'accept-encoding': 'gzip, deflate, br'}
        r = requests.get(url, headers=headers).json()
        data = r["filtered"]["data"]

    except json.decoder.JSONDecodeError:
        print('Using Dummy Data Because of json.decoder.JSONDecodeError')
        file = open('data.txt', "r")
        data = file.read()
        data = json.loads(data)
        temp = data
        data = temp["filtered"]["data"]
    x = data

    list_data = []

    for i in range(len(x)):
        stkp = x[i]['strikePrice']
        # print(stkp)
        if stkp >= 9300:
            list_data.append(x[i])

    len_list_data = len(list_data)

    list_List = []
    api_list = []
    for i in range(len_list_data):
        CE_data = list_data[i]['CE']
        stkp = CE_data['strikePrice']
        ce_oi = CE_data['openInterest'] * 75
        ce_chng_in_oi = CE_data['changeinOpenInterest'] * 75
        ce_volume = CE_data['totalTradedVolume']
        ce_ltp = CE_data['lastPrice']
        CE = {'strikePrice': stkp, 'openInterest': ce_oi / 75, 'changeInOpenInterest': ce_chng_in_oi / 75,
              'volume': ce_volume, 'lastPrice': ce_ltp}
        api_list.append(CE)
        listing = [ce_oi, ce_chng_in_oi, ce_volume, ce_ltp, stkp]
        list_List.append(listing)

    tuple_List = tuple(list_List)
    print(tuple_List)

    row_num = 1
    col_num = 0

    # function for directory
    directory = "PyProgram"

    # Parent Directory path
    parent_dir = "C:/"

    # calling function of Directory
    createDir(parent_dir, directory)

    # print("Today date is: ", today)
    directory = str(today)

    # calling function of Directory
    createDir('C:/PyProgram/', directory)

    directory = cTimeH

    parent_dir = 'C:/PyProgram/' + str(today) + '/'

    # calling function of Directory
    createDir(parent_dir, directory)

    directory = cTime

    parent_dir = 'C:/PyProgram/' + str(today) + '/' + cTimeH + '/'

    # calling function of Directory
    createDir(parent_dir, directory)

    # Final Directory Working
    f_dir = parent_dir + directory + '/'
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(f_dir + 'Data.xlsx')
    worksheet = workbook.add_worksheet()

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': 1})

    # Write some data headers.
    worksheet.write('R1C1', 'OI', bold)
    worksheet.write('R1C2', 'CHNG IN OI', bold)
    worksheet.write('R1C3', 'VOLUME', bold)
    worksheet.write('R1C4', 'LTP', bold)
    worksheet.write('R1C4', 'STKP', bold)

    for oi, cio, vol, ltp, stkp in tuple_List:
        worksheet.write(row_num, col_num, oi)
        worksheet.write(row_num, col_num + 1, cio)
        worksheet.write(row_num, col_num + 2, vol)
        worksheet.write(row_num, col_num + 3, ltp)
        worksheet.write(row_num, col_num + 4, stkp)
        row_num = row_num + 1

    row_num = 2

    list_List2 = []
    api_list2 = []
    for i in range(len_list_data):
        PE_data = list_data[i]['PE']
        stkp = PE_data['strikePrice']
        pe_oi = PE_data['openInterest'] * 75
        pe_chng_in_oi = PE_data['changeinOpenInterest'] * 75
        pe_volume = PE_data['totalTradedVolume']
        pe_ltp = PE_data['lastPrice']
        row_num1 = row_num + i
        d_oi = '=I' + str(row_num1) + '-A' + str(row_num1)
        d_chng_oi = '=H' + str(row_num1) + '-B' + str(row_num1)
        PE = {'strikePrice': stkp, 'openInterest': pe_oi / 75, 'changeInOpenInterest': pe_chng_in_oi / 75,
              'volume': pe_volume, 'lastPrice': pe_ltp}
        api_list2.append(PE)
        listing = [pe_oi, pe_chng_in_oi, pe_volume, pe_ltp, d_oi, d_chng_oi]
        list_List2.append(listing)

    tuple_List2 = tuple(list_List2)

    row_num = 1

    # Write some other data headers.
    worksheet.write('I1', 'OI', bold)
    worksheet.write('H1', 'CHNG IN OI', bold)
    worksheet.write('G1', 'VOLUME', bold)
    worksheet.write('F1', 'LTP', bold)
    worksheet.write('J1', 'DIFF OI', bold)
    worksheet.write('K1', 'DIFF CHNG OI', bold)

    for oi, cio, vol, ltp, doi, dcoi in tuple_List2:
        worksheet.write(row_num, col_num + 8, oi)
        worksheet.write(row_num, col_num + 7, cio)
        worksheet.write(row_num, col_num + 6, vol)
        worksheet.write(row_num, col_num + 5, ltp)
        worksheet.write(row_num, col_num + 9, doi)
        worksheet.write(row_num, col_num + 10, dcoi)
        row_num = row_num + 1
    workbook.close()

    apni_api = {'CE': api_list, 'PE': api_list2}
    file = open('apni_api_data.txt', "w")
    apni_api_data = file.write(str(apni_api))
    print(apni_api)

    app = flask.Flask(__name__)
    app.config["DEBUG"] = True

    # Create some test data for our catalog in the form of a list of dictionaries.
    books = apni_api


    @app.route('/', methods=['GET'])
    def api_all():
        return jsonify(books)


    app.run()
    time.sleep(296)
