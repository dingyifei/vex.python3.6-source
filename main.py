# coding=utf-8

import json
import os
import pprint
import time
import xlwt
import matplotlib.pyplot as plt
from decimal import getcontext, Decimal
from urllib.request import urlopen

# from math import pi
# import webbrowser
# import errno
# import numpy as np
# import seaborn as sns
# import pandas as pd


# preload

getcontext().prec = 6

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("#Cover", cell_overwrite_ok=True)
sheet2 = book.add_sheet("#Matches", cell_overwrite_ok=True)
sheet3 = book.add_sheet("#Important Data", cell_overwrite_ok=True)
sheet4 = book.add_sheet("#Blank", cell_overwrite_ok=True)
sheet5 = book.add_sheet("#For World", cell_overwrite_ok=True)
sheet6 = book.add_sheet("#What We Need", cell_overwrite_ok=True)
sheet7 = book.add_sheet("#Team Spot 1", cell_overwrite_ok=True)
sheet8 = book.add_sheet("#Team Spot 2", cell_overwrite_ok=True)
sheet9 = book.add_sheet("#Team Spot 3", cell_overwrite_ok=True)
sheet10 = book.add_sheet("#Team Spot 4", cell_overwrite_ok=True)
sheet11 = book.add_sheet("#Bugged Teams", cell_overwrite_ok=True)

time_now = "Last Update:" + time.strftime("%c")
sheet1.write(2, 1, time_now)
sheet1.write(3, 1,
             "Because of there are no data for these teams: 1119S, 7386A, 8000X, 8000Z, 19771B, 30638A, 36632A, "
             "37073A, 60900A, 76921B, 99556A, 99691E, 99691H are not include in the sheet #Important Data")
STYLE_1 = xlwt.easyxf('pattern: pattern solid, fore_colour red;''font: colour white, bold True;')
STYLE_2 = xlwt.easyxf('pattern: pattern solid, fore_colour blue;''font: colour white, bold True;')
STYLE_3 = xlwt.easyxf('pattern: pattern solid, fore_colour pink;''font: colour white, bold True;')
STYLE_4 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;''font: colour white, bold True;')
STYLE_RED = xlwt.easyxf('font: colour red, bold True;')
STYLE_BLUE = xlwt.easyxf('font: colour blue, bold True;')
STYLE_BLACK = xlwt.easyxf('pattern: pattern solid, fore_colour black;''font: colour white, bold True;')
STYLE_B = xlwt.easyxf('font: colour black, bold True;')
STYLE_70 = xlwt.easyxf('pattern: pattern solid, fore_colour red;''font: colour white, bold True;')
STYLE_50 = xlwt.easyxf('pattern: pattern solid, fore_colour light_orange;''font: colour white, bold True;')
STYLE_30 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;''font: colour white, bold True;')
STYLE_0 = xlwt.easyxf('pattern: pattern solid, fore_colour bright_green;''font: colour black, bold True;')

sheet2.write(0, 0, "Team")
sheet2.write(0, 1, "Wins")
sheet2.write(0, 2, "Losses")
sheet2.write(0, 3, "AP")
sheet2.write(0, 4, "Ranking")
sheet2.write(0, 5, "Highest")
sheet2.write(0, 6, "Result")


class GlobalVar:
    # used in graphbubble, graphred, timeisout
    teamr1 = ""
    teamr2 = ""
    teamr3 = ""

    # used in graphbubble, graphblue, timeisout
    teamb1 = ""
    teamb2 = ""
    teamb3 = ""


    # only used in teamskill and timeisout
    teamname = ""

    # used in graphbubble, graphred, timeisout
    teamr1wins = 0
    teamr2wins = 0
    teamr3wins = 0

    # used in graphbubble, graphblue, timeisout
    teamb1wins = 0
    teamb2wins = 0
    teamb3wins = 0

    # used in teamskill and timeisout
    skillave = 0

    # used in answer, graphbubble, graphred, timeisout
    teamr1skillout = 0
    teamr2skillout = 0
    teamr3skillout = 0
    teamb1skillout = 0
    teamb2skillout = 0
    teamb3skillout = 0

    # only graphbubble, graphred, and timeisout
    teamr1ap = 0
    teamr2ap = 0
    teamr3ap = 0

    # only graphbubble, graphblue, and timeisout
    teamb1ap = 0
    teamb2ap = 0
    teamb3ap = 0

    # only graphbubble, graphred, and timeisout
    teamr1ranking = 0
    teamr2ranking = 0
    teamr3ranking = 0

    # only graphbubble, graphblue, and timeisout
    teamb1ranking = 0
    teamb2ranking = 0
    teamb3ranking = 0

    # only graphbubble, graphred, and timeisout
    teamr1highest = 0
    teamr2highest = 0
    teamr3highest = 0

    # only graphbubble, graphblue, and timeisout
    teamb1highest = 0
    teamb2highest = 0
    teamb3highest = 0

    # only graphbubble and timeisout
    teamr1ccwm = 0
    teamr2ccwm = 0
    teamr3ccwm = 0
    teamb1ccwm = 0
    teamb2ccwm = 0
    teamb3ccwm = 0

    # only graphbubble and timeisout
    teamr1opr = 0
    teamr2opr = 0
    teamr3opr = 0
    teamb1opr = 0
    teamb2opr = 0
    teamb3opr = 0

    # only graphbubble and timeisout
    teamr1dpr = 0
    teamr2dpr = 0
    teamr3dpr = 0
    teamb1dpr = 0
    teamb2dpr = 0
    teamb3dpr = 0

    # Only teamcurrent and timeisout
    currentranking = 0
    currentwins = 0
    currentlosses = 0

    # only graphbubble and timeisout
    teamr1currentranking = 0
    teamr2currentranking = 0
    teamr3currentranking = 0
    teamb1currentranking = 0
    teamb2currentranking = 0
    teamb3currentranking = 0

    # only graphbubble and timeisout
    teamr1currentwins = 0
    teamr2currentwins = 0
    teamr3currentwins = 0
    teamb1currentwins = 0
    teamb2currentwins = 0
    teamb3currentwins = 0
    teamr1currentlosses = 0
    teamr2currentlosses = 0
    teamr3currentlosses = 0
    teamb1currentlosses = 0
    teamb2currentlosses = 0
    teamb3currentlosses = 0

    # the crap I don't want to locate
    winsave = 0
    apave = 0
    oprave = 0
    oprtotal = 0
    dprave = 0
    rankave = 0
    highestave = 0
    ccwmave = 0


def vexdb_json(api_type: str, api_parameters: dict, return_data = None):

    """
    It function accept a string "api_type" and a dictionary "api_parameters", the "api_type" should be
    one from _API_TYPE The dictionary's key are the _parameters from vexdb.io/the_data and the value should
    also follow it.
    """
    # TODO(Yifei): Multi thread, timeout retry,throw error correctly

    if return_data is None:
        return_data = ["full"]
    _parameters = ""
    output = None

    if api_parameters:
        if type(api_parameters) == dict:
            _keys = list(api_parameters.keys())
            _values = list(api_parameters.values())
            if len(_keys) >= 1:
                _parameters += "?" + _keys[0] + "=" + _values[0]
                if len(_keys) > 1:
                    for x in range(1, len(_keys)):
                        _parameters += "&" + _keys[x] + "=" + _values[x]
    else:
        _parameters = None

    if api_type != "":
        if _parameters != "" or _parameters is not None:
            json_dict = json.loads((urlopen("https://api.vexdb.io/v1/get_" + api_type + _parameters)).read())
            if json_dict["status"] == 0:
                raise() # TODO: a exception
            else:
                if json_dict["size"] == 5000:
                    raise() # TODO: Another exception or use some trick to prevent 5000 limit
                else:
                    if return_data[0] == "full":
                        output = json_dict
                    if return_data[0] != "full":
                        output = []
                        for x in range(0, len(json_dict["result"])):
                            for y in range(0, len(return_data)):
                                output.append(json_dict["result"][x][return_data[y]])
                return output

def team_list():
    # TODO(YIFEI): MAKE IT WORK
    print(vexdb_json("teams", {"grade": "High School"},["number"]))
    print(vexdb_json("matches", {"season":"Starstruck", "team":"8667A"}, ["sku"]))


def scan_team_matches(name: object) -> object:  # TODO: temperory
    _json_dict = vexdb_json("matches", {"season": "Turning%20Point", "team": name})
    output = []
    for r in _json_dict["result"]:
        line = '{}: Match{} Round{} || Red Alliance 1 = {} Red Alliance 2 = {} Red Alliance 3 = {} Red Sit = {} || ' \
               'Blue Alliance 1 = {} Blue Alliance 2 = {} Blue Alliance 3 = {} Blue Sit = {} || Red Score = {} Blue ' \
               'Score = {}'.format(r["sku"], r["matchnum"], r["round"], r["red1"], r["red2"], r["red3"], r["redsit"],
                                   r["blue1"], r["blue2"], r["blue3"], r["bluesit"], r["redscore"], r["bluescore"])
        output.append(line)
    return output


def excel_scan_teams(teams: list, season: str):  # 201

    start = time.time()
    number = 0
    sheet_line = 0

    while True:
        while number < len(teams):
            teamloop = teams[number]
            # ['sheet_%d' % sheetnb].write = book.add_sheet(teamloop, cell_overwrite_ok= True)
            print('')
            print(teamloop)
            print('')
            number += 1
            sheet_line += 1
            json_dict = vexdb_json("rankings", {"team": teamloop, "season": season})
            output = []
            for r in json_dict["result"]:
                line = 'Team = {} Wins = {} Losses = {} AP = {} Ranking in Current Match = {} Highest Score = {}' \
                    .format(r["team"], r["wins"], r["losses"], r["ap"], r["rank"], r["max_score"])
                datateam = '{}'.format(r["team"])
                datawins = '{}'.format(r["wins"])
                datalosses = '{}'.format(r["losses"])
                dataap = '{}'.format(r["ap"])
                datarank = '{}'.format(r["rank"])
                datamaxscore = '{}'.format(r["max_score"])
                if int(datawins) > int(datalosses):
                    sheet2.write(sheet_line, 6, "Positive", STYLE_1)
                elif int(datawins) < int(datalosses):
                    sheet2.write(sheet_line, 6, "Negative", STYLE_2)
                output.append(line)
                # ['sheet' + str(number)].write(1, 1,teamloop)
                # ['sheet' + str(number)].write(2, 1,line )
                sheet2.write(sheet_line, 0, datateam)
                sheet2.write(sheet_line, 1, datawins)
                sheet2.write(sheet_line, 2, datalosses)
                sheet2.write(sheet_line, 3, dataap)
                sheet2.write(sheet_line, 4, datarank)
                sheet2.write(sheet_line, 5, datamaxscore)
                sheet_line += 1
            # pprint.pprint(output)
            book.save("Data.xls")
            print('')
            decimal = (time.time() - start)
            decimal = Decimal.from_float(decimal).quantize(Decimal('0.0'))
            ave = (float(decimal) / (int(number)))
            ave = Decimal.from_float(ave).quantize(Decimal('0.0'))
            eta = (float(ave) * (int(len(teams) - (int(number)))))
            etatomin = (float(eta) / 60)
            etatomin = Decimal.from_float(etatomin).quantize(Decimal('0.0'))
            print(str(number) + "/" + str(len(teams)) + " Finished, Used " + str(decimal) + " seconds. Average " + str(
                ave) + " seconds each. ETA: " + str(etatomin) + " mins.")
        if number >= 5:
            number = 0
            sheet_line = 1
            print('\n reset and xls saved!')


def excel_get_all_data(teams: list, season: str):  # 203
    sheetline = 0
    start = time.time()

    # TODO(Yifei):This List is removed NEEDFIX
    # list1 = []

    while True:
        for number in range(0, len(teams)):
            teamloop = teams[number]
            print(teamloop)
            number += 1
            teaminfoline = int(sheetline)
            sheet3.write(sheetline, 0, "Team")
            sheet3.write(sheetline, 1, "Wins")
            sheet3.write(sheetline, 2, "Losses")
            sheet3.write(sheetline, 3, "AP")
            sheet3.write(sheetline, 4, "Ranking")
            sheet3.write(sheetline, 5, "Highest")
            sheet3.write(sheetline, 6, "Result")
            sheet3.write(sheetline, 8, "Flag")
            sheetline += 1
            json_dict = vexdb_json("ranking", {"team": teamloop, "season": season})
            output = []
            for r in json_dict["result"]:
                line = "Team = {} Wins = {} Losses = {} AP = {} Ranking in Current Match = {} Highest Score = {}" \
                    .format(r["team"], r["wins"], r["losses"], r["ap"], r["rank"], r["max_score"])
                output.append(line)

            datateam = '{}'.format(r["team"])
            datawins = '{}'.format(r["wins"])
            datalosses = '{}'.format(r["losses"])
            dataap = '{}'.format(r["ap"])
            datarank = '{}'.format(r["rank"])
            datamaxscore = '{}'.format(r["max_score"])
            output.append(line)

            sheet3.write(sheetline, 0, "#" + datateam)
            sheet3.write(sheetline, 1, datawins)
            sheet3.write(sheetline, 2, datalosses)
            sheet3.write(sheetline, 3, dataap)
            sheet3.write(sheetline, 4, datarank)
            sheet3.write(sheetline, 5, datamaxscore)

            if int(datawins) > int(datalosses):
                sheet3.write(sheetline, 6, "Positive", STYLE_1)
            elif int(datawins) < int(datalosses):
                sheet3.write(sheetline, 6, "Negative", STYLE_2)
            sheetline += 1
            # pprint.pprint(output)
            json_dict = vexdb_json("matches", {"team": teamloop, "season": season})
            # print('\n')
            output = []
            loop = -10000
            # 1-10000 For testing, should be 0

            sheet3.write(sheetline, 0, "Sku")
            sheet3.write(sheetline, 1, "Match")
            sheet3.write(sheetline, 2, "Red1")
            sheet3.write(sheetline, 3, "Red2")
            sheet3.write(sheetline, 4, "Red3")
            sheet3.write(sheetline, 5, "RedSit")
            sheet3.write(sheetline, 6, "Blue1")
            sheet3.write(sheetline, 7, "Blue2")
            sheet3.write(sheetline, 8, "Blue3")
            sheet3.write(sheetline, 9, "BlueSit")
            sheet3.write(sheetline, 10, "RedSco")
            sheet3.write(sheetline, 11, "BlueSco")
            sheet3.write(sheetline, 12, "Team LF")
            sheet3.write(sheetline, 13, "Result")
            sheet3.write(sheetline, 14, "Difficulty")
            sheet3.write(sheetline, 15, "Status")
            # sheet3.write(sheetline, 16, "Cur Sit")

            sheetline += 1
            win = 0
            matches = 0
            for r in json_dict["result"]:
                matches += 1
                line = '{}: Match{} Round{} || Red Alliance 1 = {} Red Alliance 2 = {} Red Alliance 3 = {} Red Sit = ' \
                       '{} || Blue Alliance 1 = {} Blue Alliance 2 = {} Blue Alliance 3 = {} Blue Sit = {} || Red ' \
                       'Score = {} Blue Score = {}'.format(r["sku"], r["matchnum"], r["round"], r["red1"], r["red2"],
                                                           r["red3"], r["redsit"], r["blue1"], r["blue2"], r["blue3"],
                                                           r["bluesit"], r["redscore"], r["bluescore"])
                datasku = '{}'.format(r["sku"])
                datamatchnum = '{}'.format(r["matchnum"])
                datared1 = '{}'.format(r["red1"])
                datared2 = '{}'.format(r["red2"])
                datared3 = '{}'.format(r["red3"])
                dataredsit = '{}'.format(r["redsit"])
                datablue1 = '{}'.format(r["blue1"])
                datablue2 = '{}'.format(r["blue2"])
                datablue3 = '{}'.format(r["blue3"])
                databluesit = '{}'.format(r["bluesit"])
                dataredsc = '{}'.format(r["redscore"])
                databluesc = '{}'.format(r["bluescore"])
                # sheetline += 1
                sheet3.write(sheetline, 0, datasku)
                sheet3.write(sheetline, 1, datamatchnum)
                sheet3.write(sheetline, 2, datared1, STYLE_RED)
                sheet3.write(sheetline, 3, datared2, STYLE_RED)
                sheet3.write(sheetline, 4, datared3, STYLE_RED)
                sheet3.write(sheetline, 5, dataredsit, STYLE_RED)
                sheet3.write(sheetline, 6, datablue1, STYLE_BLUE)
                sheet3.write(sheetline, 7, datablue2, STYLE_BLUE)
                sheet3.write(sheetline, 8, datablue3, STYLE_BLUE)
                sheet3.write(sheetline, 9, databluesit, STYLE_BLUE)
                sheet3.write(sheetline, 10, dataredsc, STYLE_RED)
                sheet3.write(sheetline, 11, databluesc, STYLE_BLUE)
                sheet3.write(sheetline, 12, datateam + " =", STYLE_B)

                if int(dataredsc) > int(databluesc):
                    sheet3.write(sheetline, 14, "Red", STYLE_1)
                elif int(dataredsc) < int(databluesc):
                    sheet3.write(sheetline, 14, "Blue", STYLE_2)

                if int(dataredsc) + 20 < int(databluesc):
                    sheet3.write(sheetline, 14, "Blue Easy", STYLE_4)
                elif int(dataredsc) - 20 > int(databluesc):
                    sheet3.write(sheetline, 14, "Red Easy", STYLE_3)

                if datared1 == teamloop or datared2 == teamloop or datared3 == teamloop:
                    if int(dataredsc) > int(databluesc):
                        sheet3.write(sheetline, 13, "Win", STYLE_B)
                        win += 1
                    else:
                        sheet3.write(sheetline, 13, "Lose", STYLE_BLACK)

                elif datablue1 == teamloop or datablue2 == teamloop or datablue3 == teamloop:
                    if int(dataredsc) < int(databluesc):
                        sheet3.write(sheetline, 13, "Win", STYLE_B)
                        win += 1
                    else:
                        sheet3.write(sheetline, 13, "Lose", STYLE_BLACK)

                # To see if 0 = 0
                if int(dataredsc) == 0 and int(databluesc) == 0:
                    sheetline -= 1
                    matches -= 1

                elif int(dataredsc) == 0:
                    sheet3.write(sheetline, 15, "Red DQ?", STYLE_BLACK)
                elif int(databluesc) == 0:
                    sheet3.write(sheetline, 15, "Blue DQ?", STYLE_BLACK)

                sheetline += 1
                loop += 1

                if loop > 2:
                    break

                output.append(line)

            sheetline += 1

            teaminfoline += 1

            decimal = (int(win) / int(matches))
            flag = decimal * 100
            flag = Decimal.from_float(flag).quantize(Decimal('0.0'))

            if float(flag) >= 70:
                sheet3.write(teaminfoline, 8, str(flag) + "%", STYLE_70)
                for x in range(9, 21):
                    sheet3.write(teaminfoline, x, "", STYLE_70)

            elif float(flag) >= 50:
                sheet3.write(teaminfoline, 8, str(flag) + "%", STYLE_50)
                for x in range(9, 21):
                    sheet3.write(teaminfoline, x, "", STYLE_50)

            elif float(flag) >= 30:
                sheet3.write(teaminfoline, 8, str(flag) + "%", STYLE_30)
                for x in range(9, 21):
                    sheet3.write(teaminfoline, x, "", STYLE_30)

            else:
                sheet3.write(teaminfoline, 8, str(flag) + "%", STYLE_0)
                for x in range(9, 21):
                    sheet3.write(teaminfoline, x, "", STYLE_0)
            for x in range(0, 21):
                sheet3.write(sheetline, x, "- - - - - - -", STYLE_BLACK)

            sheetline += 1

            decimal = (time.time() - start)
            decimal = Decimal.from_float(decimal).quantize(Decimal('0.0'))

            ave = (float(decimal) / (int(number)))
            ave = Decimal.from_float(ave).quantize(Decimal('0.0'))

            eta = (float(ave) * (int(len(teams) - (int(number)))))
            etatomin = (float(eta) / 60)
            etatomin = Decimal.from_float(etatomin).quantize(Decimal('0.0'))

            print(str(number) + "/" + str(len(teams)) + " Finished, Used " + str(decimal) + " seconds. Average " + str(
                ave) + " seconds each. ETA: " + str(etatomin) + " mins.")
            print()
            book.save("Data" + ".xls")

        if number >= 5:
            number = 0
            sheetline = 1
            print('')
            print('reset and xls saved!')


def excel_get_all_bugs(teams: list, season: str):  # 204

    number = 0
    sheetline = 0
    start = time.time()

    while True:
        while number < int(len(teams)):
            teamloop = teams[number]
            print(teamloop)
            number += 1
            teaminfoline = int(sheetline)
            sheet10.write(sheetline, 0, "Team")
            sheet10.write(sheetline, 1, "Wins")
            sheet10.write(sheetline, 2, "Losses")
            sheet10.write(sheetline, 3, "AP")
            sheet10.write(sheetline, 4, "Ranking")
            sheet10.write(sheetline, 5, "Highest")
            sheet10.write(sheetline, 6, "Result")
            sheet10.write(sheetline, 8, "Flag")
            sheetline += 1
            json_dict = vexdb_json("rankings", {"team": teamloop, "season": season})
            output = []

            for r in json_dict["result"]:
                line = "Team = {} Wins = {} Losses = {} AP = {} Ranking in Current Match = {} Highest Score = {}" \
                    .format(r["team"], r["wins"], r["losses"], r["ap"], r["rank"], r["max_score"])
                output.append(line)

            datateam = '{}'.format(r["team"])
            datawins = '{}'.format(r["wins"])
            datalosses = '{}'.format(r["losses"])
            dataap = '{}'.format(r["ap"])
            datarank = '{}'.format(r["rank"])
            datamaxscore = '{}'.format(r["max_score"])
            output.append(line)

            sheet10.write(sheetline, 0, "#" + datateam)
            sheet10.write(sheetline, 1, datawins)
            sheet10.write(sheetline, 2, datalosses)
            sheet10.write(sheetline, 3, dataap)
            sheet10.write(sheetline, 4, datarank)
            sheet10.write(sheetline, 5, datamaxscore)

            if int(datawins) > int(datalosses):
                sheet10.write(sheetline, 6, "Positive", STYLE_1)
            elif int(datawins) < int(datalosses):
                sheet10.write(sheetline, 6, "Negative", STYLE_2)

            sheetline += 1

            json_dict = vexdb_json("matches", {"team": teamloop, "season": season})
            output = []

            loop = -10000

            # 1-10000 For testing, should be 0

            sheet10.write(sheetline, 0, "Sku")
            sheet10.write(sheetline, 1, "Match")
            sheet10.write(sheetline, 2, "Red1")
            sheet10.write(sheetline, 3, "Red2")
            sheet10.write(sheetline, 4, "Red3")
            sheet10.write(sheetline, 5, "RedSit")
            sheet10.write(sheetline, 6, "Blue1")
            sheet10.write(sheetline, 7, "Blue2")
            sheet10.write(sheetline, 8, "Blue3")
            sheet10.write(sheetline, 9, "BlueSit")
            sheet10.write(sheetline, 10, "RedSco")
            sheet10.write(sheetline, 11, "BlueSco")
            sheet10.write(sheetline, 12, "Team LF")
            sheet10.write(sheetline, 13, "Result")
            sheet10.write(sheetline, 14, "Difficulty")
            sheet10.write(sheetline, 15, "Status")
            # sheet10.write(sheetline, 16, "Cur Sit")

            sheetline += 1

            win = 0
            matches = 0

            for r in json_dict["result"]:

                matches += 1

                line = '{}: Match{} Round{} || Red Alliance 1 = {} Red Alliance 2 = {} Red Alliance 3 = {} Red Sit = ' \
                       '{} || Blue Alliance 1 = {} Blue Alliance 2 = {} Blue Alliance 3 = {} Blue Sit = {} || Red ' \
                       'Score = {} Blue Score = {}' \
                    .format(r["sku"], r["matchnum"], r["round"], r["red1"], r["red2"], r["red3"], r["redsit"],
                            r["blue1"], r["blue2"], r["blue3"], r["bluesit"], r["redscore"], r["bluescore"])

                datasku = '{}'.format(r["sku"])
                datamatchnum = '{}'.format(r["matchnum"])
                datared1 = '{}'.format(r["red1"])
                datared2 = '{}'.format(r["red2"])
                datared3 = '{}'.format(r["red3"])
                dataredsit = '{}'.format(r["redsit"])
                datablue1 = '{}'.format(r["blue1"])
                datablue2 = '{}'.format(r["blue2"])
                datablue3 = '{}'.format(r["blue3"])
                databluesit = '{}'.format(r["bluesit"])
                dataredsc = '{}'.format(r["redscore"])
                databluesc = '{}'.format(r["bluescore"])

                sheet10.write(sheetline, 0, datasku)
                sheet10.write(sheetline, 1, datamatchnum)
                sheet10.write(sheetline, 2, datared1, STYLE_RED)
                sheet10.write(sheetline, 3, datared2, STYLE_RED)
                sheet10.write(sheetline, 4, datared3, STYLE_RED)
                sheet10.write(sheetline, 5, dataredsit, STYLE_RED)
                sheet10.write(sheetline, 6, datablue1, STYLE_BLUE)
                sheet10.write(sheetline, 7, datablue2, STYLE_BLUE)
                sheet10.write(sheetline, 8, datablue3, STYLE_BLUE)
                sheet10.write(sheetline, 9, databluesit, STYLE_BLUE)
                sheet10.write(sheetline, 10, dataredsc, STYLE_RED)
                sheet10.write(sheetline, 11, databluesc, STYLE_BLUE)
                sheet10.write(sheetline, 12, datateam + " =", STYLE_B)

                if int(dataredsc) > int(databluesc):
                    sheet10.write(sheetline, 14, "Red", STYLE_1)
                elif int(dataredsc) < int(databluesc):
                    sheet10.write(sheetline, 14, "Blue", STYLE_2)

                if int(dataredsc) + 20 < int(databluesc):
                    sheet10.write(sheetline, 14, "Blue Easy", STYLE_4)
                elif int(dataredsc) - 20 > int(databluesc):
                    sheet10.write(sheetline, 14, "Red Easy", STYLE_3)

                if datared1 == teamloop or datared2 == teamloop or datared3 == teamloop:
                    if int(dataredsc) > int(databluesc):
                        sheet10.write(sheetline, 13, "Win", STYLE_B)
                        win += 1
                    else:
                        sheet10.write(sheetline, 13, "Lose", STYLE_BLACK)

                elif datablue1 == teamloop or datablue2 == teamloop or datablue3 == teamloop:
                    if int(dataredsc) < int(databluesc):
                        sheet10.write(sheetline, 13, "Win", STYLE_B)
                        win += 1
                    else:
                        sheet10.write(sheetline, 13, "Lose", STYLE_BLACK)

                # To see if 0 = 0

                if int(dataredsc) == 0 and int(databluesc) == 0:
                    sheetline -= 1
                    matches -= 1

                elif int(dataredsc) == 0:
                    sheet10.write(sheetline, 15, "Red DQ?", STYLE_BLACK)
                elif int(databluesc) == 0:
                    sheet10.write(sheetline, 15, "Blue DQ?", STYLE_BLACK)

                sheetline += 1
                loop += 1

                if loop > 2:
                    break

                output.append(line)

            sheetline += 1

            teaminfoline += 1

            decimal = (int(win) / int(matches))
            flag = decimal * 100
            flag = Decimal.from_float(flag).quantize(Decimal('0.0'))

            if float(flag) >= 70:
                sheet10.write(teaminfoline, 8, str(flag) + "%", STYLE_70)
                for x in range(9, 21):
                    sheet10.write(teaminfoline, x, "", STYLE_70)

            elif float(flag) >= 50:
                sheet10.write(teaminfoline, 8, str(flag) + "%", STYLE_50)
                for x in range(9, 21):
                    sheet10.write(teaminfoline, x, "", STYLE_50)

            elif float(flag) >= 30:
                sheet10.write(teaminfoline, 8, str(flag) + "%", STYLE_30)
                for x in range(9, 21):
                    sheet10.write(teaminfoline, x, "", STYLE_30)

            else:
                sheet10.write(teaminfoline, 8, str(flag) + "%", STYLE_0)
                for x in range(9, 21):
                    sheet10.write(teaminfoline, x, "", STYLE_0)
            for x in range(0, 21):
                sheet10.write(sheetline, x, "- - - - - - -", STYLE_BLACK)

            sheetline += 1

            decimal = (time.time() - start)
            decimal = Decimal.from_float(decimal).quantize(Decimal('0.0'))

            ave = (float(decimal) / (int(number)))
            ave = Decimal.from_float(ave).quantize(Decimal('0.0'))

            eta = float(ave) * (int(len(teams) - (int(number))))
            etatomin = (float(eta) / 60)
            etatomin = Decimal.from_float(etatomin).quantize(Decimal('0.0'))

            print(str(number) + "/" + str(len(teams)) + " Finished, Used " + str(decimal) + " seconds. Average " + str(
                ave) + " seconds each. ETA: " + str(etatomin) + " mins.")
            print()
            book.save("Data" + ".xls")

        if number >= 5:
            number = 0
            sheetline = 1
            print('')
            print('reset and xls saved!')


def excel_get_we_need(teams: list, season: str):  # 205

    number = 0
    sheetline = 0
    start = time.time()

    while True:  # Todo(Yifei): What is this loop for?
        while number < int(len(teams)):  # TODO(Yifei): Use for loop instead

            teamloop = teams[number]
            print(teamloop)
            number += 1
            teaminfoline = int(sheetline)
            sheet6.write(sheetline, 0, "Team")
            sheet6.write(sheetline, 1, "Wins")
            sheet6.write(sheetline, 2, "Losses")
            sheet6.write(sheetline, 3, "AP")
            sheet6.write(sheetline, 4, "Ranking")
            sheet6.write(sheetline, 5, "Highest")
            sheet6.write(sheetline, 6, "Result")
            sheet6.write(sheetline, 8, "Flag")
            sheetline += 1
            json_dict = vexdb_json("rankings", {"team": teamloop, "season": season})
            output = []
            for r in json_dict["result"]:
                line = "Team = {} Wins = {} Losses = {} AP = {} Ranking in Current Match = {} Highest Score = {}" \
                    .format(r["team"], r["wins"], r["losses"], r["ap"], r["rank"], r["max_score"])
                output.append(line)
            datateam = '{}'.format(r["team"])
            datawins = '{}'.format(r["wins"])
            datalosses = '{}'.format(r["losses"])
            dataap = '{}'.format(r["ap"])
            datarank = '{}'.format(r["rank"])
            datamaxscore = '{}'.format(r["max_score"])
            output.append(line)
            sheet6.write(sheetline, 0, "#" + datateam)
            sheet6.write(sheetline, 1, datawins)
            sheet6.write(sheetline, 2, datalosses)
            sheet6.write(sheetline, 3, dataap)
            sheet6.write(sheetline, 4, datarank)
            sheet6.write(sheetline, 5, datamaxscore)

            if int(datawins) > int(datalosses):
                sheet6.write(sheetline, 6, "Positive", STYLE_1)
            elif int(datawins) < int(datalosses):
                sheet6.write(sheetline, 6, "Negative", STYLE_2)

            sheetline += 1

            json_dict = vexdb_json("matches", {"team": teamloop, "season": season})
            output = []
            loop = -10000
            # 1-10000 For testing, should be 0

            sheet6.write(sheetline, 0, "Sku")
            sheet6.write(sheetline, 1, "Match")
            sheet6.write(sheetline, 2, "Red1")
            sheet6.write(sheetline, 3, "Red2")
            sheet6.write(sheetline, 4, "Red3")
            sheet6.write(sheetline, 5, "RedSit")
            sheet6.write(sheetline, 6, "Blue1")
            sheet6.write(sheetline, 7, "Blue2")
            sheet6.write(sheetline, 8, "Blue3")
            sheet6.write(sheetline, 9, "BlueSit")
            sheet6.write(sheetline, 10, "RedSco")
            sheet6.write(sheetline, 11, "BlueSco")
            sheet6.write(sheetline, 12, "Team LF")
            sheet6.write(sheetline, 13, "Result")
            sheet6.write(sheetline, 14, "Difficulty")
            sheet6.write(sheetline, 15, "Status")
            # sheet6.write(sheetline, 16, "Difference")

            sheetline += 1

            win = 0
            matches = 0

            for r in json_dict["result"]:

                matches += 1

                line = '{}: Match{} Round{} || Red Alliance 1 = {} Red Alliance 2 = {} Red Alliance 3 = {} Red Sit = ' \
                       '{} || Blue Alliance 1 = {} Blue Alliance 2 = {} Blue Alliance 3 = {} Blue Sit = {} || Red ' \
                       'Score = {} Blue Score = {}' \
                    .format(r["sku"], r["matchnum"], r["round"], r["red1"], r["red2"], r["red3"], r["redsit"],
                            r["blue1"], r["blue2"], r["blue3"], r["bluesit"], r["redscore"], r["bluescore"])

                datasku = '{}'.format(r["sku"])
                datamatchnum = '{}'.format(r["matchnum"])
                datared1 = '{}'.format(r["red1"])
                datared2 = '{}'.format(r["red2"])
                datared3 = '{}'.format(r["red3"])
                dataredsit = '{}'.format(r["redsit"])
                datablue1 = '{}'.format(r["blue1"])
                datablue2 = '{}'.format(r["blue2"])
                datablue3 = '{}'.format(r["blue3"])
                databluesit = '{}'.format(r["bluesit"])
                dataredsc = '{}'.format(r["redscore"])
                databluesc = '{}'.format(r["bluescore"])

                # sheetline += 1

                sheet6.write(sheetline, 0, datasku)
                sheet6.write(sheetline, 1, datamatchnum)
                sheet6.write(sheetline, 2, datared1, STYLE_RED)
                sheet6.write(sheetline, 3, datared2, STYLE_RED)
                sheet6.write(sheetline, 4, datared3, STYLE_RED)
                sheet6.write(sheetline, 5, dataredsit, STYLE_RED)
                sheet6.write(sheetline, 6, datablue1, STYLE_BLUE)
                sheet6.write(sheetline, 7, datablue2, STYLE_BLUE)
                sheet6.write(sheetline, 8, datablue3, STYLE_BLUE)
                sheet6.write(sheetline, 9, databluesit, STYLE_BLUE)
                sheet6.write(sheetline, 10, dataredsc, STYLE_RED)
                sheet6.write(sheetline, 11, databluesc, STYLE_BLUE)
                sheet6.write(sheetline, 12, datateam + " =", STYLE_B)

                if int(dataredsc) > int(databluesc):
                    sheet6.write(sheetline, 14, "Red", STYLE_1)
                elif int(dataredsc) < int(databluesc):
                    sheet6.write(sheetline, 14, "Blue", STYLE_2)

                if int(dataredsc) + 20 < int(databluesc):
                    sheet6.write(sheetline, 14, "Blue Easy", STYLE_4)
                elif int(dataredsc) - 20 > int(databluesc):
                    sheet6.write(sheetline, 14, "Red Easy", STYLE_3)

                if datared1 == teamloop or datared2 == teamloop or datared3 == teamloop:
                    if int(dataredsc) > int(databluesc):
                        sheet6.write(sheetline, 13, "Win", STYLE_B)
                        win += 1
                    else:
                        sheet6.write(sheetline, 13, "Lose", STYLE_BLACK)

                elif datablue1 == teamloop or datablue2 == teamloop or datablue3 == teamloop:
                    if int(dataredsc) < int(databluesc):
                        sheet6.write(sheetline, 13, "Win", STYLE_B)
                        win += 1
                    else:
                        sheet6.write(sheetline, 13, "Lose", STYLE_BLACK)

                # To see if 0 = 0

                if int(dataredsc) == 0 and int(databluesc) == 0:
                    sheetline -= 1
                    matches -= 1
                elif int(dataredsc) == 0:
                    sheet6.write(sheetline, 15, "Red DQ?", STYLE_BLACK)
                elif int(databluesc) == 0:
                    sheet6.write(sheetline, 15, "Blue DQ?", STYLE_BLACK)

                sheetline += 1
                loop += 1

                if loop > 2:
                    break

                output.append(line)

            sheetline += 1
            teaminfoline += 1

            decimal = (int(win) / int(matches))
            flag = decimal * 100
            flag = Decimal.from_float(flag).quantize(Decimal('0.0'))

            if float(flag) >= 70:
                sheet6.write(teaminfoline, 8, str(flag) + "%", STYLE_70)
                for x in range(9, 21):
                    sheet6.write(teaminfoline, x, "", STYLE_70)
            elif float(flag) >= 50:
                sheet6.write(teaminfoline, 8, str(flag) + "%", STYLE_50)
                for x in range(9, 21):
                    sheet6.write(teaminfoline, x, "", STYLE_50)
            elif float(flag) >= 30:
                sheet6.write(teaminfoline, 8, str(flag) + "%", STYLE_30)
                for x in range(9, 21):
                    sheet6.write(teaminfoline, x, "", STYLE_30)
            else:
                sheet6.write(teaminfoline, 8, str(flag) + "%", STYLE_0)
                for x in range(9, 21):
                    sheet6.write(teaminfoline, x, "", STYLE_0)

            for x in range(0, 21):
                sheet6.write(sheetline, x, "- - - - - - -", STYLE_BLACK)

            sheetline += 1

            decimal = (time.time() - start)
            decimal = Decimal.from_float(decimal).quantize(Decimal('0.0'))

            ave = (float(decimal) / (int(number)))
            ave = Decimal.from_float(ave).quantize(Decimal('0.0'))

            eta = float(ave) * (int(len(teams) - (int(number))))
            etatomin = (float(eta) / 60)
            etatomin = Decimal.from_float(etatomin).quantize(Decimal('0.0'))

            print(str(number) + "/" + str(len(teams)) + " Finished, Used " + str(decimal) + " seconds. Average " + str(
                ave) + " seconds each. ETA: " + str(etatomin) + " mins.")
            print()
            book.save("Data" + ".xls")

        if number >= 5:
            number = 0
            sheetline = 1
            print('\n reset and xls saved!')


def excel_scan_world(teams: list, season: str, sku: str):
    number = 0
    sheetline = 0
    start = time.time()
    while True:
        while number < int(len(teams)):
            teamloop = teams[number]
            print(teamloop)
            number += 1
            sheet5.write(sheetline, 0, "Team")
            sheet5.write(sheetline, 1, "Wins")
            sheet5.write(sheetline, 2, "Losses")
            sheet5.write(sheetline, 3, "AP")
            sheet5.write(sheetline, 4, "Ranking")
            sheet5.write(sheetline, 5, "Highest")
            sheet5.write(sheetline, 6, "Result")
            sheetline += 1
            json_dict = vexdb_json("rankings", {"team": teamloop, "season": season, "sku": sku})
            output = []

            for r in json_dict["result"]:
                line = "Team = {} Wins = {} Losses = {} AP = {} Ranking in Current Match = {} Highest Score = {}" \
                    .format(r["team"], r["wins"], r["losses"], r["ap"], r["rank"], r["max_score"])
                output.append(line)

            datateam = '{}'.format(r["team"])
            datawins = '{}'.format(r["wins"])
            datalosses = '{}'.format(r["losses"])
            dataap = '{}'.format(r["ap"])
            datarank = '{}'.format(r["rank"])
            datamaxscore = '{}'.format(r["max_score"])

            # output.append(line) #Remove because I cant see the use of this

            sheet5.write(sheetline, 0, "#" + datateam)
            sheet5.write(sheetline, 1, datawins)
            sheet5.write(sheetline, 2, datalosses)
            sheet5.write(sheetline, 3, dataap)
            sheet5.write(sheetline, 4, datarank)
            sheet5.write(sheetline, 5, datamaxscore)

            if int(datawins) > int(datalosses):
                sheet5.write(sheetline, 6, "Positive", STYLE_1)
            elif int(datawins) < int(datalosses):
                sheet5.write(sheetline, 6, "Negative", STYLE_2)

            sheetline += 1
            json_dict = vexdb_json("matches", {"team": teamloop, "season": season})
            output = []
            loop = -10000

            sheet5.write(sheetline, 0, "Sku")
            sheet5.write(sheetline, 1, "Match")
            sheet5.write(sheetline, 2, "Red1")
            sheet5.write(sheetline, 3, "Red2")
            sheet5.write(sheetline, 4, "Red3")
            sheet5.write(sheetline, 5, "RedSit")
            sheet5.write(sheetline, 6, "Blue1")
            sheet5.write(sheetline, 7, "Blue2")
            sheet5.write(sheetline, 8, "Blue3")
            sheet5.write(sheetline, 9, "BlueSit")
            sheet5.write(sheetline, 10, "RedSco")
            sheet5.write(sheetline, 11, "BlueSco")

            for r in json_dict["result"]:
                line = '{}: Match{} Round{} || Red Alliance 1 = {} Red Alliance 2 = {} Red Alliance 3 = {} Red Sit = ' \
                       '{} || Blue Alliance 1 = {} Blue Alliance 2 = {} Blue Alliance 3 = {} Blue Sit = {} || Red ' \
                       'Score = {} Blue Score = {}' \
                    .format(r["sku"], r["matchnum"], r["round"], r["red1"], r["red2"], r["red3"], r["redsit"],
                            r["blue1"], r["blue2"], r["blue3"], r["bluesit"], r["redscore"], r["bluescore"])
                datasku = '{}'.format(r["sku"])
                datamatchnum = '{}'.format(r["matchnum"])
                datared1 = '{}'.format(r["red1"])
                datared2 = '{}'.format(r["red2"])
                datared3 = '{}'.format(r["red3"])
                dataredsit = '{}'.format(r["redsit"])
                datablue1 = '{}'.format(r["blue1"])
                datablue2 = '{}'.format(r["blue2"])
                datablue3 = '{}'.format(r["blue3"])
                databluesit = '{}'.format(r["bluesit"])
                dataredsc = '{}'.format(r["redscore"])
                databluesc = '{}'.format(r["bluescore"])

                sheetline += 1

                sheet5.write(sheetline, 0, datasku)
                sheet5.write(sheetline, 1, datamatchnum)
                sheet5.write(sheetline, 2, datared1, STYLE_RED)
                sheet5.write(sheetline, 3, datared2, STYLE_RED)
                sheet5.write(sheetline, 4, datared3, STYLE_RED)
                sheet5.write(sheetline, 5, dataredsit, STYLE_RED)
                sheet5.write(sheetline, 6, datablue1, STYLE_BLUE)
                sheet5.write(sheetline, 7, datablue2, STYLE_BLUE)
                sheet5.write(sheetline, 8, datablue3, STYLE_BLUE)
                sheet5.write(sheetline, 9, databluesit, STYLE_BLUE)
                sheet5.write(sheetline, 10, dataredsc, STYLE_RED)
                sheet5.write(sheetline, 11, databluesc, STYLE_BLUE)
                sheet5.write(sheetline, 12, datateam + " =", STYLE_B)

                if int(dataredsc) > int(databluesc):
                    sheet5.write(sheetline, 14, "Red", STYLE_1)
                elif int(dataredsc) < int(databluesc):
                    sheet5.write(sheetline, 14, "Blue", STYLE_2)

                if int(dataredsc) + 20 < int(databluesc):
                    sheet5.write(sheetline, 14, "Blue Easy", STYLE_4)
                elif int(dataredsc) - 20 > int(databluesc):
                    sheet5.write(sheetline, 14, "Red Easy", STYLE_3)

                if datared1 == teamloop or datared2 == teamloop or datared3 == teamloop:
                    if int(dataredsc) > int(databluesc):
                        sheet5.write(sheetline, 13, "Win", STYLE_B)
                    else:
                        sheet5.write(sheetline, 13, "Lose", STYLE_BLACK)
                elif datablue1 == teamloop or datablue2 == teamloop or datablue3 == teamloop:
                    if int(dataredsc) < int(databluesc):
                        sheet5.write(sheetline, 13, "Win", STYLE_B)
                    else:
                        sheet5.write(sheetline, 13, "Lose", STYLE_BLACK)

                sheetline += 1
                loop += 1

                if loop > 2:
                    break

                output.append(line)
            sheetline += 1
            for x in range(0, 15):
                sheet5.write(sheetline, x, "- - - - - - -", STYLE_BLACK)

            sheetline += 1
            for x in range(0, 15):
                sheet5.write(sheetline, x, "- - - - - - -", STYLE_BLACK)
            sheetline += 1

            decimal = (time.time() - start)
            decimal = Decimal.from_float(decimal).quantize(Decimal('0.0'))

            ave = (float(decimal) / (int(number)))
            ave = Decimal.from_float(ave).quantize(Decimal('0.0'))

            eta = float(ave) * (int(len(teams) - (int(number))))
            etatomin = (float(eta) / 60)
            etatomin = Decimal.from_float(etatomin).quantize(Decimal('0.0'))

            print(str(number) + "/" + str(len(teams)) + " Finished, Used " + str(decimal) + " seconds. Average " + str(
                ave) + " seconds each. ETA: " + str(etatomin) + " mins.")
            print()
            book.save("Data" + ".xls")

        if number >= 5:
            number = 0
            sheetline = 1
            print('\n reset and xls saved!')
            main()


# Need to test when competition start


def excel_team_matches(name, season):  # TODO(YIFEI): Why excel? Value name change.

    _json_dict = vexdb_json("matches", {"team": name, "season": season})
    output = []
    for r in _json_dict["result"]:
        line = '{}: Match{} Round{} || Red Alliance 1 = {} Red Alliance 2 = {} Red Alliance 3 = {} Red Sit = {} || ' \
               'Blue Alliance 1 = {} Blue Alliance 2 = {} Blue Alliance 3 = {} Blue Sit = {} || Red Score = {} Blue ' \
               'Score = {}' \
            .format(r["sku"], r["matchnum"], r["round"], r["red1"], r["red2"], r["red3"], r["redsit"], r["blue1"],
                    r["blue2"], r["blue3"], r["bluesit"], r["redscore"], r["bluescore"])
        output.append(line)
    return output


def search_team_current_season(name, season):  # TODO(YIFEI): Value name change.

    json_dict = vexdb_json("rankings", {"team": name, "season": season})
    output = []
    for r in json_dict["result"]:
        line = "Team = {} Wins = {} Losses = {} AP = {} Ranking in Current Match = {} Highest Score = {}" \
            .format(r["team"], r["wins"], r["losses"], r["ap"], r["rank"], r["max_score"])
        output.append(line)
    return output


def get_all_data(name, season):

    # print("This will show the recent three matches.")
    json_dict = vexdb_json("ranking", {"team": name, "season": season})
    ranking_result = []
    for r in json_dict["result"]:
        line = "Team = {} Wins = {} Losses = {} AP = {} Ranking in Current Match = {} Highest Score = {}" \
            .format(r["team"], r["wins"], r["losses"], r["ap"], r["rank"], r["max_score"])
        ranking_result.append(line)
    json_dict = vexdb_json("matches", {"team": name, "season": season})
    matches_result = []
    loop = 0
    for r in json_dict["result"]:
        line = '{}: Match{} Round{} || Red Alliance 1 = {} Red Alliance 2 = {} Red Alliance 3 = {} Red Sit = {} || ' \
               'Blue Alliance 1 = {} Blue Alliance 2 = {} Blue Alliance 3 = {} Blue Sit = {} || Red Score = {} Blue ' \
               'Score = {}' \
            .format(r["sku"], r["matchnum"], r["round"], r["red1"], r["red2"], r["red3"], r["redsit"], r["blue1"],
                    r["blue2"], r["blue3"], r["bluesit"], r["redscore"], r["bluescore"])
        loop += 1
        if loop > 2:
            break
        matches_result.append(line)
    return ranking_result, matches_result


def time_is_out(red_teams: list, blue_teams: list, season: str):  # TODO: NEED MORE FIX

    # GlobalVar.inputmode = str(input("Type in the preset value or 6 teams separate by ,\n"))
    for x in range(0, len(red_teams) - 1):  #TODO(YIFEI): Make it work
        result = team_skill(red_teams[x], season)


    for x in range(0, len(blue_teams) - 1):
        result = team_skill(blue_teams[x], season)



    if str(GlobalVar.teamr1) != "":
        GlobalVar.teamname = GlobalVar.teamr1
        team_skill()
        GlobalVar.teamr1skillout = GlobalVar.skillave
        GlobalVar.teamr1wins = GlobalVar.winsave
        GlobalVar.teamr1ap = GlobalVar.apave
        GlobalVar.teamr1ranking = GlobalVar.rankave
        GlobalVar.teamr1highest = GlobalVar.highestave
        GlobalVar.teamr1ccwm = GlobalVar.ccwmave
        GlobalVar.teamr1dpr = GlobalVar.dprave
        GlobalVar.teamr1opr = GlobalVar.oprave
        GlobalVar.teamr1currentranking = GlobalVar.currentranking
        GlobalVar.teamr1currentwins = GlobalVar.currentwins
        GlobalVar.teamr1currentlosses = GlobalVar.currentlosses
    else:
        print("Team Red 1 is blank.")

    if str(GlobalVar.teamr2) != "":
        GlobalVar.teamname = GlobalVar.teamr2
        team_skill()
        GlobalVar.teamr2skillout = GlobalVar.skillave
        GlobalVar.teamr2wins = GlobalVar.winsave
        GlobalVar.teamr2ap = GlobalVar.apave
        GlobalVar.teamr2ranking = GlobalVar.rankave
        GlobalVar.teamr2highest = GlobalVar.highestave
        GlobalVar.teamr2ccwm = GlobalVar.ccwmave
        GlobalVar.teamr2dpr = GlobalVar.dprave
        GlobalVar.teamr2opr = GlobalVar.oprave
        GlobalVar.teamr2currentranking = GlobalVar.currentranking
        GlobalVar.teamr2currentwins = GlobalVar.currentwins
        GlobalVar.teamr2currentlosses = GlobalVar.currentlosses
    else:
        print("Team Red 2 is blank.")

    if str(GlobalVar.teamr3) != "":
        GlobalVar.teamname = GlobalVar.teamr3
        team_skill()
        GlobalVar.teamr3skillout = GlobalVar.skillave
        GlobalVar.teamr3wins = GlobalVar.winsave
        GlobalVar.teamr3ap = GlobalVar.apave
        GlobalVar.teamr3ranking = GlobalVar.rankave
        GlobalVar.teamr3highest = GlobalVar.highestave
        GlobalVar.teamr3ccwm = GlobalVar.ccwmave
        GlobalVar.teamr3dpr = GlobalVar.dprave
        GlobalVar.teamr3opr = GlobalVar.oprave
        GlobalVar.teamr3currentranking = GlobalVar.currentranking
        GlobalVar.teamr3currentwins = GlobalVar.currentwins
        GlobalVar.teamr3currentlosses = GlobalVar.currentlosses
    else:
        print("Team Red 3 is blank.")

    if str(GlobalVar.teamb1) != "":
        GlobalVar.teamname = GlobalVar.teamb1
        team_skill()
        GlobalVar.teamb1skillout = GlobalVar.skillave
        GlobalVar.teamb1wins = GlobalVar.winsave
        GlobalVar.teamb1ap = GlobalVar.apave
        GlobalVar.teamb1ranking = GlobalVar.rankave
        GlobalVar.teamb1highest = GlobalVar.highestave
        GlobalVar.teamb1ccwm = GlobalVar.ccwmave
        GlobalVar.teamb1dpr = GlobalVar.dprave
        GlobalVar.teamb1opr = GlobalVar.oprave
        GlobalVar.teamb1currentranking = GlobalVar.currentranking
        GlobalVar.teamb1currentwins = GlobalVar.currentwins
        GlobalVar.teamb1currentlosses = GlobalVar.currentlosses
    else:
        print("Team Blue 1 is blank.")

    if str(GlobalVar.teamb2) != "":
        GlobalVar.teamname = GlobalVar.teamb2
        team_skill()
        GlobalVar.teamb2skillout = GlobalVar.skillave
        GlobalVar.teamb2wins = GlobalVar.winsave
        GlobalVar.teamb2ap = GlobalVar.apave
        GlobalVar.teamb2ranking = GlobalVar.rankave
        GlobalVar.teamb2highest = GlobalVar.highestave
        GlobalVar.teamb2ccwm = GlobalVar.ccwmave
        GlobalVar.teamb2dpr = GlobalVar.dprave
        GlobalVar.teamb2opr = GlobalVar.oprave
        GlobalVar.teamb2currentranking = GlobalVar.currentranking
        GlobalVar.teamb2currentwins = GlobalVar.currentwins
        GlobalVar.teamb2currentlosses = GlobalVar.currentlosses
    else:
        print("Team Blue 2 is blank.")

    if str(GlobalVar.teamb3) != "":
        GlobalVar.teamname = GlobalVar.teamb3
        team_skill()
        GlobalVar.teamb3skillout = GlobalVar.skillave
        GlobalVar.teamr3wins = GlobalVar.winsave
        GlobalVar.teamb3ap = GlobalVar.apave
        GlobalVar.teamb3ranking = GlobalVar.rankave
        GlobalVar.teamb3highest = GlobalVar.highestave
        GlobalVar.teamb3ccwm = GlobalVar.ccwmave
        GlobalVar.teamb3dpr = GlobalVar.dprave
        GlobalVar.teamb3opr = GlobalVar.oprave
        GlobalVar.teamb3currentranking = GlobalVar.currentranking
        GlobalVar.teamb3currentwins = GlobalVar.currentwins
        GlobalVar.teamb3currentlosses = GlobalVar.currentlosses
    else:
        print("Team Blue 3 is blank.")

    graphbubble()  # pass value use arg instead of global

    return None


def team_skill(team, season):

    json_dict = vexdb_json("skills", {"team": team, "season": season})  # it should be globalvar teamsent
    skilltotal = 0
    totalattempts = 0
    for r in json_dict["result"]:
        skill = int(r["score"])
        attempt = int(r["attempts"])
        if int(attempt) != 0:
            totalattempts += 1
        skilltotal += skill

    if int(totalattempts) != 0:
        skillave = int(skilltotal) / int(totalattempts)
    else:
        skillave = 0
    decimal = skillave
    decimal = Decimal.from_float(decimal).quantize(Decimal('0.0'))
    GlobalVar.skillave = decimal
    print(GlobalVar.teamname + ": " + str(GlobalVar.skillave))
    team_sent()


def team_sent(team, season):

    count = 0
    GlobalVar.winsave = 0
    json_dict = vexdb_json("rankings", {"season": season, "team": team})  # it should be teamsent
    for r in json_dict["result"]:
        teamwins = '{}'.format(r["wins"])
        count += 1
        winstotal = teamwins + teamwins
        if teamwins == "" or teamwins == "":
            print("break cuz blank")
            count -= 1
            GlobalVar.winsave = float(winstotal) / int(count)
            teamap()
        GlobalVar.winsave = float(winstotal) / int(count)
    team_current()


def team_current(team, season, sku):  # can be part of teamsent()

    GlobalVar.currentranking = 0
    GlobalVar.currentwins = 0
    GlobalVar.currentlosses = 0
    json_dict = vexdb_json("rankings", {"season": season, "team": team, "sku": sku})  # teamsent, sku constant, season
    for r in json_dict["result"]:
        GlobalVar.currentranking = '{}'.format(r["rank"])
        GlobalVar.currentwins = '{}'.format(r["wins"])
        GlobalVar.currentlosses = '{}'.format(r["losses"])
    teamap()


def teamap(team,season):

    aptotal = 0
    count = 0
    json_dict = vexdb_json("rankings", {"team":team, "season":season})  #teamsent
    for r in json_dict["result"]:
        teammap = '{}'.format(r["ap"])
        count += 1
        if int(teammap) > 25:
            diff = (int(teammap) - 25) * 0.2
            teammap = 25 + float(diff)
            print("Balance over 25, " + str(diff))
        aptotal += int(teammap)
        GlobalVar.apave = int(aptotal) / int(count)
        if teammap == "" or teammap == "":
            print("break cuz blank")
            count -= 1
            teamranking()
    teamranking()


def teamranking(team, season):
    GlobalVar.rankave = 0
    count = 0
    rank_total = 0
    json_dict = vexdb_json("rankings", {"team": team, "season": season})  # teamsent
    for r in json_dict["result"]:
        team_ranking = '{}'.format(r["rank"])
        count += 1
        rank_total += int(team_ranking)
        GlobalVar.rankave = float(rank_total) / count
        if team_ranking == "":
            print("break cuz blank")
            count -= 1
            GlobalVar.rankave = float(rank_total) / count
            team_highest()
        GlobalVar.rankave = float(team_ranking) / count
    team_highest()


def team_highest(team, season):

    highesttotal = 0
    GlobalVar.highestave = 0
    count = 0
    json_dict = vexdb_json("rankings", {"team": team, "season": season})  # teamsent
    for r in json_dict["result"]:
        team_highest = '{}'.format(r["max_score"])
        count += 1
        highesttotal += int(team_highest)
        GlobalVar.highestave = int(highesttotal) / count
        if team_highest == "":
            print("break cuz blank")
            count -= 1
            GlobalVar.highestave = float(highesttotal) / count
            teampr()
        GlobalVar.highestave = float(highesttotal) / count
    teampr()


def teampr(team, season):

    GlobalVar.oprtotal = 0
    dprtotal = 0
    json_dict = vexdb_json("rankings", {"team": team, "season": season})  #teamsent
    count = 0
    for r in json_dict["result"]:
        teamopr = '{}'.format(r["opr"])
        teamdpr = '{}'.format(r["dpr"])
        teamopr = (float(teamopr) / 5)
        teamdpr = (float(teamdpr) / 5)
        count += 1
        GlobalVar.oprtotal += float(teamopr)
        GlobalVar.oprave = float(GlobalVar.oprtotal) / count
        dprtotal += float(teamdpr)
        GlobalVar.dprave = float(dprtotal) / count
        if teamdpr == "" or teamopr == "":
            print("break cuz blank")
            count -= 1
            teamccwm()

        teamccwm()


def teamccwm(team, season):

    ccwmtotal = 0
    GlobalVar.ccwmave = 0
    json_dict = vexdb_json("rankings", {"team":team, "season": season}) # teamsent
    count = 0
    for r in json_dict["result"]:
        teamccwm = '{}'.format(r["ccwm"])
        count += 1
        ccwmtotal += float(teamccwm)
        GlobalVar.ccwmave = float(ccwmtotal) / count
        if teamccwm == "" or teamccwm == "":
            print("break cuz blank")
            count -= 18
            break


def graphbubble(file_name: str):  # it should be part of "timeisout"
    GlobalVar.teamr1skillout = float(GlobalVar.teamr1skillout) / 10
    GlobalVar.teamr2skillout = float(GlobalVar.teamr2skillout) / 10
    GlobalVar.teamr3skillout = float(GlobalVar.teamr3skillout) / 10
    GlobalVar.teamr1ap = round(float(GlobalVar.teamr1ap) / 5, 1)
    GlobalVar.teamr2ap = round(float(GlobalVar.teamr2ap) / 5, 1)
    GlobalVar.teamr3ap = round(float(GlobalVar.teamr3ap) / 5, 1)
    # The Formula
    GlobalVar.teamr1ranking = int(10 - int(GlobalVar.teamr1ranking))
    GlobalVar.teamr2ranking = int(10 - int(GlobalVar.teamr2ranking))
    GlobalVar.teamr3ranking = int(10 - int(GlobalVar.teamr3ranking))

    # /17
    GlobalVar.teamr1highest = round(float(int(GlobalVar.teamr1highest) / 17), 1)
    GlobalVar.teamr2highest = round(float(int(GlobalVar.teamr2highest) / 17), 1)
    GlobalVar.teamr3highest = round(float(int(GlobalVar.teamr3highest) / 17), 1)

    if int(GlobalVar.teamr1ranking) < 0:
        GlobalVar.teamr1ranking = 0
    if int(GlobalVar.teamr2ranking) < 0:
        GlobalVar.teamr2ranking = 0
    if int(GlobalVar.teamr3ranking) < 0:
        GlobalVar.teamr3ranking = 0

    # Check
    print("Skill " + str(GlobalVar.teamr1skillout) + " " + str(GlobalVar.teamr2skillout) + " " + str(
        GlobalVar.teamr3skillout))
    print("Season Wins " + str(GlobalVar.teamr1wins) + " " + str(GlobalVar.teamr2wins) + " " + str(
        GlobalVar.teamr3wins))
    print("AP " + str(GlobalVar.teamr1ap) + " " +
          str(GlobalVar.teamr2ap) + " " + str(GlobalVar.teamr3ap))
    print("Ranking " + str(GlobalVar.teamr1ranking) + " " + str(GlobalVar.teamr2ranking) + " " + str(
        GlobalVar.teamr3ranking))
    print("Highest " + str(GlobalVar.teamr1highest) + " " + str(GlobalVar.teamr2highest) + " " + str(
        GlobalVar.teamr3highest))
    print("CCWM" + str(GlobalVar.teamr1ccwm))

    GlobalVar.teamb1skillout = float(GlobalVar.teamb1skillout) / 10
    GlobalVar.teamb2skillout = float(GlobalVar.teamb2skillout) / 10
    GlobalVar.teamb3skillout = float(GlobalVar.teamb3skillout) / 10

    GlobalVar.teamb1ap = round(float(GlobalVar.teamb1ap) / 5, 1)
    GlobalVar.teamb2ap = round(float(GlobalVar.teamb2ap) / 5, 1)
    GlobalVar.teamb3ap = round(float(GlobalVar.teamb3ap) / 5, 1)

    # The Formula
    GlobalVar.teamb1ranking = int(10 - int(GlobalVar.teamb1ranking))
    GlobalVar.teamb2ranking = int(10 - int(GlobalVar.teamb2ranking))
    GlobalVar.teamb3ranking = int(10 - int(GlobalVar.teamb3ranking))

    # /17
    GlobalVar.teamb1highest = round(
        float(int(GlobalVar.teamb1highest) / 17), 1)
    GlobalVar.teamb2highest = round(
        float(int(GlobalVar.teamb2highest) / 17), 1)
    GlobalVar.teamb3highest = round(
        float(int(GlobalVar.teamb3highest) / 17), 1)

    if int(GlobalVar.teamb1ranking) <= 0:
        GlobalVar.teamb1ranking = 0
    if int(GlobalVar.teamb2ranking) <= 0:
        GlobalVar.teamb2ranking = 0
    if int(GlobalVar.teamb3ranking) <= 0:
        GlobalVar.teamb3ranking = 0

    # Check
    print("Skill " + str(GlobalVar.teamb1skillout) + " " + str(GlobalVar.teamb2skillout) + " " + str(
        GlobalVar.teamb3skillout))
    print("Season Wins " + str(GlobalVar.teamb1wins) + " " + str(GlobalVar.teamb2wins) + " " + str(
        GlobalVar.teamb3wins))
    print("AP " + str(GlobalVar.teamb1ap) + " " +
          str(GlobalVar.teamb2ap) + " " + str(GlobalVar.teamb3ap))
    print("Ranking " + str(GlobalVar.teamb1ranking) + " " + str(GlobalVar.teamb2ranking) + " " + str(
        GlobalVar.teamb3ranking))
    print("Highest " + str(GlobalVar.teamb1highest) + " " + str(GlobalVar.teamb2highest) + " " + str(
        GlobalVar.teamb3highest))

    if GlobalVar.teamr1ccwm < 0:
        GlobalVar.teamr1ccwm = 0.1
    if GlobalVar.teamr2ccwm < 0:
        GlobalVar.teamr2ccwm = 0.1
    if GlobalVar.teamr3ccwm < 0:
        GlobalVar.teamr3ccwm = 0.1
    if GlobalVar.teamb1ccwm < 0:
        GlobalVar.teamb1ccwm = 0.1
    if GlobalVar.teamb2ccwm < 0:
        GlobalVar.teamb2ccwm = 0.1
    if GlobalVar.teamb3ccwm < 0:
        GlobalVar.teamb3ccwm = 0.1

    # create data!

    x = float(GlobalVar.teamr1skillout)
    y = float(GlobalVar.teamr1ap)
    # z = float(GlobalVar.teamr1wins)
    z = float(GlobalVar.teamr1highest)
    plt.text(x, y, str(GlobalVar.teamr1), ha='center',
             va='center', fontweight='bold', color='red')
    plt.scatter(x, y, s=z * 300, c="red", alpha=0.4, linewidth=6)

    x = float(GlobalVar.teamr2skillout)
    y = float(GlobalVar.teamr2ap)
    # z = float(GlobalVar.teamr2wins)
    z = float(GlobalVar.teamr2highest)
    plt.text(x, y, str(GlobalVar.teamr2), ha='center',
             va='center', fontweight='bold', color='red')
    plt.scatter(x, y, s=z * 300, c="red", alpha=0.4, linewidth=6)

    x = float(GlobalVar.teamr3skillout)
    y = float(GlobalVar.teamr3ap)
    # z = float(GlobalVar.teamr3wins)
    z = float(GlobalVar.teamr3highest)
    plt.text(x, y, str(GlobalVar.teamr3), ha='center',
             va='center', fontweight='bold', color='red')
    plt.scatter(x, y, s=z * 300, c="red", alpha=0.4, linewidth=6)

    x = float(GlobalVar.teamr1dpr)
    y = float(GlobalVar.teamr1opr)
    # z = float(GlobalVar.teamr1wins)
    z = float(GlobalVar.teamr1ccwm)
    plt.text(x, y, str("[" + GlobalVar.teamr1 + "]"), ha='center',
             fontweight='bold', va='center', color='darkred')
    plt.scatter(x, y, s=z * 50, c="deeppink", alpha=0.4, linewidth=6)

    x = float(GlobalVar.teamr2dpr)
    y = float(GlobalVar.teamr2opr)
    # z = float(GlobalVar.teamr2wins)
    z = float(GlobalVar.teamr2ccwm)
    plt.text(x, y, str("[" + GlobalVar.teamr2 + "]"), ha='center',
             fontweight='bold', va='center', color='darkred')
    plt.scatter(x, y, s=z * 50, c="deeppink", alpha=0.4, linewidth=6)

    if GlobalVar.teamr3dpr != 0:
        x = float(GlobalVar.teamr3dpr)
        y = float(GlobalVar.teamr3opr)
        # z = float(GlobalVar.teamr3wins)
        z = float(GlobalVar.teamr3ccwm)
        plt.text(x, y, str("[" + GlobalVar.teamr3 + "]"), ha='center',
                 fontweight='bold', va='center', color='darkred')
        plt.scatter(x, y, s=z * 50, c="deeppink", alpha=0.4, linewidth=6)

    x = float(GlobalVar.teamb1skillout)
    y = float(GlobalVar.teamb1ap)
    # z = float(GlobalVar.teamb1wins)
    z = float(GlobalVar.teamb1highest)
    plt.text(x, y, str(GlobalVar.teamb1), ha='center',
             va='center', fontweight='bold', color='royalblue')
    plt.scatter(x, y, s=z * 300, c="royalblue", alpha=0.4, linewidth=6)

    x = float(GlobalVar.teamb2skillout)
    y = float(GlobalVar.teamb2ap)
    # z = float(GlobalVar.teamb2wins)
    z = float(GlobalVar.teamb2highest)
    plt.text(x, y, str(GlobalVar.teamb2), ha='center',
             va='center', fontweight='bold', color='royalblue')
    plt.scatter(x, y, s=z * 300, c="royalblue", alpha=0.4, linewidth=6)

    x = float(GlobalVar.teamb3skillout)
    y = float(GlobalVar.teamb3ap)
    # z = float(GlobalVar.teamb3wins)
    z = float(GlobalVar.teamb3highest)
    plt.text(x, y, str(GlobalVar.teamb3), ha='center',
             va='center', fontweight='bold', color='royalblue')
    plt.scatter(x, y, s=z * 300, c="royalblue", alpha=0.4, linewidth=6)

    x = float(GlobalVar.teamb1dpr)
    y = float(GlobalVar.teamb1opr)
    # z = float(GlobalVar.teamb1wins)
    z = float(GlobalVar.teamb1ccwm)
    plt.text(x, y, str("[" + GlobalVar.teamb1 + "]"), ha='center',
             va='bottom', fontweight='bold', color='dodgerblue')
    plt.scatter(x, y, s=z * 50, c="dodgerblue", alpha=0.4, linewidth=6)

    x = float(GlobalVar.teamb2dpr)
    y = float(GlobalVar.teamb2opr)
    # z = float(GlobalVar.teamb2wins)
    z = float(GlobalVar.teamb2ccwm)
    plt.text(x, y, str("[" + GlobalVar.teamb2 + "]"), ha='center',
             va='bottom', fontweight='bold', color='dodgerblue')
    plt.scatter(x, y, s=z * 50, c="dodgerblue", alpha=0.4, linewidth=6)

    if GlobalVar.teamb3dpr != 0:
        x = float(GlobalVar.teamb3dpr)
        y = float(GlobalVar.teamb3opr)
        # z = float(GlobalVar.teamb3wins)
        z = float(GlobalVar.teamb3ccwm)
        plt.text(x, y, str("[" + GlobalVar.teamb3 + "]"), ha='center', va='bottom', fontweight='bold',
                 color='dodgerblue')
        plt.scatter(x, y, s=z * 50, c="dodgerblue", alpha=0.4, linewidth=6)

    xmin, xmax = plt.xlim()
    ymin, ymax = plt.ylim()
    xaxis = float(xmax)
    xmiddle = (float(xaxis) / 2)
    # Add titles (main and on axis)
    try: # TODO(YIFEI): It should raise error instead, remove without notice is BAD, also this try sucks
        os.remove("graph/" + file_name + ".png")
        print("Previous deleted.")
    except OSError:
        print("something is not right")
        pass
    plt.xlabel("Skill / [Defensive]")
    plt.ylabel("AP / [Offensive]")
    plt.title(
        "Red: " + GlobalVar.teamr1 + " " + GlobalVar.teamr2 + " " + GlobalVar.teamr3 +
        " Blue: " + GlobalVar.teamb1 + " " + GlobalVar.teamb2 + " " + GlobalVar.teamb3,loc="left")
    plt.text(xmiddle, -0.02,
             "Team #, X: Skill, Y: AP, Z: Highest Score\n [Team #], X: Defensive Pts Y: Offensive Pts Z: Contribution",
             ha='center', color='white', bbox=dict(facecolor='darkslateblue', alpha=0.5))
    plt.text((xmin + 0.3), (ymax - 0.5), GlobalVar.teamr1 + " W: " + str(GlobalVar.teamr1currentwins) + " L: " + str(
        GlobalVar.teamr1currentlosses) + " R: " + str(
        GlobalVar.teamr1currentranking) + "\n" + GlobalVar.teamr2 + " W: " + str(
        GlobalVar.teamr2currentwins) + " L: " + str(GlobalVar.teamr2currentlosses) + " R: " + str(
        GlobalVar.teamr2currentranking) + "\n" + GlobalVar.teamr3 + " W: " + str(
        GlobalVar.teamr3currentwins) + " L: " + str(GlobalVar.teamr3currentlosses) + " R: " + str(
        GlobalVar.teamr3currentranking) + "\n" + GlobalVar.teamb1 + " W: " + str(
        GlobalVar.teamb1currentwins) + " L: " + str(GlobalVar.teamb1currentlosses) + " R: " + str(
        GlobalVar.teamb1currentranking) + "\n" + GlobalVar.teamb2 + " W: " + str(
        GlobalVar.teamb2currentwins) + " L: " + str(GlobalVar.teamb2currentlosses) + " R: " + str(
        GlobalVar.teamb2currentranking) + "\n" + GlobalVar.teamb3 + " W: " + str(
        GlobalVar.teamb3currentwins) + " L: " + str(GlobalVar.teamb3currentlosses) + " R: " + str(
        GlobalVar.teamb3currentranking), ha='left', va='top', color='white', fontsize='smaller',
             bbox=dict(facecolor='darkgreen', alpha=0.5))
    plt.savefig("graph/" + file_name + ".png")
    print("Graph poped and saved.")
    plt.show()


def answer():

    teamrexist = 0
    teambexist = 0

    if GlobalVar.teamr1skillout != 0:
        teamrexist += 1
    if GlobalVar.teamr2skillout != 0:
        teamrexist += 1
    if GlobalVar.teamr3skillout != 0:
        teamrexist += 1
    if GlobalVar.teamb1skillout != 0:
        teambexist += 1
    if GlobalVar.teamb2skillout != 0:
        teambexist += 1
    if GlobalVar.teamb3skillout != 0:
        teambexist += 1
    input("Press Any Key to Continue\n")


def main():
    while True:
        mode = int(input(  #TODO(YIFEI): int??? exception
            "Mode \n 1.Scan Team Matches \n 2.Excel Functions [Not Finished] \n 3.Search Team Season History  "
            "\n 8.Get Important Info For a Team \n 9.Change Log\n 0.Quit \n"))
        if mode == 1:
            print("Mode = Scan Team Matches")
            print(scan_team_matches(input("team number:")))
        elif mode == 2:
            print("Mode = Excels")
            print(
                "1.Scan Teams \n2.Scan Matches [Don't use this]\n3.Write Team Important Data\n4.Don't Ues This\n5.Can "
                "Specific Match [PreSet World Championship]\n6.Get We Need")
            excel_mode = int(input())
            if excel_mode == 1:
                print("Mode = Scan Teams and Write to Excel")
                excel_scan_teams()
            elif excel_mode == 2:
                print("Mode = Write Team Matches [Don't use this]")
                input1 = input('Team #?\n')
                input2 = input('season #?\n')
                pprint.pprint(excel_team_matches(input1, input2))
            elif excel_mode == 3:
                print("Mode = Write Team Important Data in Excel")
                excel_get_all_data()
            elif excel_mode == 4:
                print("Mode = Scan Bugged Team [It will crash]")
                excel_get_all_bugs()
            elif excel_mode == 5:
                print("Mode = Scan World Championship")
                excel_scan_world()
            elif excel_mode == 6:
                print("Mode = Scan We Need")
                excel_get_we_need()
        elif mode == 3:
            print("Mode = Search Team History : Current Season")
            input1 = input('Team #?\n')
            input2 = input('season #?\n')
            pprint.pprint(search_team_current_season(input1, input2))
        elif mode == 4:
            print("Bubble!")
            time_is_out()
            answer()
        elif mode == 8:
            print("Mode = Get Important Data")
            input1 = input('Team #?\n')
            input2 = input('season #?\n')
            a = get_all_data(input1,input2)
            pprint.pprint(a[0])
            pprint.pprint(a[1])
        elif mode == 0:
            print("Thanks for using it!")
            quit()

if __name__ == '__main__':
        main()
