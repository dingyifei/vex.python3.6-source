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
sleep_timer = 0
book = xlwt.Workbook(encoding="utf-8")

VEXDB_API_MATCHES = 'https://api.vexdb.io/v1/get_matches?team='
VEXDB_API_RANK = 'https://api.vexdb.io/v1/get_rankings?team='
VEX_SEASON = '&season=Turning%20Point'

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

now = time.strftime("%c")
time_now = "Last Update:" + time.strftime("%c")
sheet1.write(2, 1, time_now)
sheet1.write(3, 1,
             "Because of there are no data for these teams: 1119S, 7386A, 8000X, 8000Z, 19771B, 30638A, 36632A, "
             "37073A, 60900A, 76921B, 99556A, 99691E, 99691H are not include in the sheet #Important Data")

STYLE_1 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour red;''font: colour white, bold True;')
STYLE_2 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour blue;''font: colour white, bold True;')
STYLE_3 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour pink;''font: colour white, bold True;')
STYLE_4 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour pale_blue;''font: colour white, bold True;')
STYLE_RED = xlwt.easyxf(
    'font: colour red, bold True;')
STYLE_BLUE = xlwt.easyxf(
    'font: colour blue, bold True;')
STYLE_BLACK = xlwt.easyxf(
    'pattern: pattern solid, fore_colour black;''font: colour white, bold True;')
STYLE_B = xlwt.easyxf(
    'font: colour black, bold True;')
STYLE_70 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour red;''font: colour white, bold True;')
STYLE_50 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour light_orange;''font: colour white, bold True;')
STYLE_30 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour pale_blue;''font: colour white, bold True;')
STYLE_0 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour bright_green;''font: colour black, bold True;')

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

    # Move to public
    teamsent = ""

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
    inputmode = ""

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

    CONST_match = "&sku=RE-VRC-17-3805"  # Only used in Team current

    # the crap I don't want to locate
    winsave = 0
    apave = 0
    oprave = 0
    oprtotal = 0
    dprave = 0
    rankave = 0
    highestave = 0
    ccwmave = 0


def scan_team_matches():
    name = input('Team #?\n')
    print('Checking, TEAM %s.' % name)

    r = urlopen(VEXDB_API_MATCHES + name + VEX_SEASON)
    text = r.read()
    pprint.pprint(json.loads(text))
    json_dict = json.loads(text)
    print('\n')
    output = []

    for r in json_dict["result"]:
        line = '{}: Match{} Round{} || Red Alliance 1 = {} Red Alliance 2 = {} Red Alliance 3 = {} Red Sit = {} || ' \
               'Blue Alliance 1 = {} Blue Alliance 2 = {} Blue Alliance 3 = {} Blue Sit = {} || Red Score = {} Blue ' \
               'Score = {}'.format(r["sku"], r["matchnum"], r["round"], r["red1"], r["red2"], r["red3"], r["redsit"],
                                   r["blue1"], r["blue2"], r["blue3"], r["bluesit"], r["redscore"], r["bluescore"])
        output.append(line)

    pprint.pprint(output)
    time.sleep(1)

    return None


def excel_scan_teams():  # 201

    start = time.time()
    number = 0
    sheetline = 0
    list1 = ['2S', '2U', '5S', '10N', '12C', '12E', '12F', '12G', '12J', '39A', '39J', '39K', '39W', '39Y', '46B',
             '56C', '60X', '62A', '66A', '81K', '81Y', '91C', '109A', '114T', '127X', '134C', '134D', '134E', '134G',
             '136N', '162A', '169A', '169C', '169E', '169Y', '170A', '177V', '180A', '180X', '183Z', '185A', '202Z',
             '231X', '244A', '244B', '278E', '285X', '288A', '306X', '315G', '315J', '315X', '315Z', '321A', '323G',
             '323Y', '333R', '343X', '355D', '355E', '355M', '356A', '356B', '356C', '359A', '359X', '363A', '365X',
             '398A', '409A', '464M', '507D', '523X', '536C', '536E', '546A', '574C', '574F', '590G', '590V', '598B',
             '599A', '621A', '624H', '624K', '643T', '643Z', '660A', '666U', '666X', '675D', '686A', '709S', '815J',
             '817A', '824Z', '901C', '917F', '920C', '929H', '929U', '929X', '934Z', '970A', '986A', '1008M', '1010B',
             '1010N', '1010X', '1028A', '1028B', '1028Z', '1045A', '1045B', '1064A', '1069E', '1104S', '1115A', '1119S',
             '1138B', '1193Z', '1200Z', '1233F', '1235C', '1248B', '1264D', '1267C', '1275A', '1275B', '1275D', '1320B',
             '1320C', '1320D', '1344A', '1353C', '1356B', '1410A', '1429B', '1437Z', '1460A', '1483B', '1492W', '1492X',
             '1492Y', '1492Z', '1505R', '1533M', '1575X', '1588A', '1588D', '1617A', '1690X', '1718A', '1727E', '1727F',
             '1784Z', '1814D', '1859S', '1859W', '1859X', '1961K', '1961N', '1961U', '1961X', '1965A', '1965R', '1965T',
             '1970K', '1973B', '2011C', '2011E', '2011F', '2019B', '2019F', '2030A', '2030B', '2114X', '2114Z', '2131E',
             '2131M', '2131R', '2131W', '2131X', '2142D', '2223Z', '2235A', '2250K', '2263C', '2284A', '2284B', '2297A',
             '2316A', '2360S', '2360V', '2373A', '2396A', '2435B', '2435W', '2442B', '2442C', '2456A', '2496V', '2496Y',
             '2560E', '2560S', '2567X', '2612A', '2616D', '2616J', '2616Y', '2719B', '2719D', '2775B', '2777V', '2777W',
             '2886B', '2900B', '2900C', '2900G', '2921S', '2941A', '2979A', '2990B', '2993M', '3018A', '3018V', '3050A',
             '3050C', '3118B', '3141S', '3159A', '3260S', '3264N', '3269A', '3269B', '3273B', '3314K', '3348B', '3388X',
             '3389D', '3547A', '3553C', '3631C', '3631Z', '3682A', '3701A', '3767A', '3767X', '3796C', '3815M', '3818D',
             '3946E', '3946W', '4004X', '4057C', '4104A', '4104C', '4142A', '4147A', '4148A', '4148D', '4154A', '4154B',
             '4169J', '4253J', '4255A', '4305A', '4306A', '4318B', '4364H', '4403A', '4409A', '4410C', '4411Y', '4454A',
             '4478V', '4478X', '4549B', '4610C', '4610Z', '4621B', '4805F', '4815A', '4815B', '4828B', '4911A', '5062A',
             '5090Z', '5106C', '5139A', '5139D', '5221T', '5225A', '5245A', '5300A', '5327A', '5327B', '5327C', '5327X',
             '5408A', '5588D', '5588E', '5588R', '5691Y', '5735B', '5735K', '5776A', '5776E', '5776T', '5864B', '5956B',
             '5956G', '5999W', '6007X', '6008A', '6008D', '6008Z', '6106B', '6106C', '6109C', '6121C', '6135H', '6135W',
             '6210Z', '6277B', '6299B', '6358A', '6358C', '6403A', '6403B', '6603V', '6627A', '6627B', '6627D', '6671X',
             '6715B', '6724B', '6741A', '6822B', '6842A', '6916C', '6916E', '6916H', '6978J', '6980A', '7110A', '7110Y',
             '7110Z', '7121D', '7121E', '7209X', '7221R', '7221T', '7232X', '7258A', '7258B', '7316D', '7316F', '7323A',
             '7368W', '7386A', '7386X', '7432B', '7432C', '7447B', '7458F', '7479A', '7536B', '7546A', '7618B', '7618C',
             '7682E', '7682S', '7686A', '7700R', '7700S', '7776B', '7842B', '7842F', '7853A', '7856A', '7862B', '7862X',
             '7884B', '7884D', '7884E', '7975F', '7983V', '7984B', '8000A', '8000B', '8000C', '8000E', '8000X', '8000Z',
             '8044A', '8059A', '8059D', '8059J', '8059X', '8059Z', '8079A', '8086A', '8110B', '8110C', '8110X', '8110Z',
             '8114A', '8176B', '8192B', '8192C', '8192D', '8223A', '8232X', '8261A', '8330A', '8331A', '8349E', '8373H',
             '8387B', '8387C', '8447A', '8451D', '8452A', '8568A', '8659G', '8669A', '8675A', '8691E', '8787A', '8800X',
             '8825S', '8855B', '8861C', '8931F', '8995M', '9020A', '9031H', '9060A', '9060B', '9060C', '9080R', '9090A',
             '9144E', '9185B', '9225C', '9228A', '9343C', '9364A', '9364C', '9364D', '9364E', '9409A', '9409B', '9409C',
             '9421X', '9457B', '9545W', '9551B', '9553E', '9594C', '9605A', '9623E', '9727A', '9784B', '9823C', '9873B',
             '9922A', '9922Z', '9932B', '9932E', '9932F', '9932G', '9973A', '10173W', '10300X', '10955M', '11124R',
             '11495A', '12298A', '15376A', '16101Z', '17071B', '17090A', '18554B', '19771B', '20610A', '20785A',
             '20785B', '21246X', '21508A', '22699A', '23880B', '25461Z', '26982E', '27183R', '30638A', '34000E',
             '34000M', '34203X', '34760A', '35211C', '35960A', '36632A', '37073A', '41364A', '41998A', '43775B',
             '44244C', '48180S', '48327M', '48667A', '48778A', '49181C', '49181N', '49181T', '49181U', '49450A',
             '50505A', '51140A', '51140B', '51581A', '53999E', '55563E', '57418A', '58072A', '59990Z', '60900A',
             '61499U', '61499Y', '61601A', '62019P', '62440F', '62993A', '64000B', '64040B', '64846A', '67292A',
             '68211A', '68555A', '71303A', '72477A', '76209G', '76565A', '76921B', '77177B', '77321J', '77788J',
             '81118P', '81785K', '86000A', '86868R', '87217R', '91709A', '93063A', '96666D', '96671B', '96671X',
             '97038A', '97140A', '97301C', '97371A', '97871A', '97934U', '98177B', '98177C', '98548B', '98725B',
             '98744A', '98807A', '99000A', '99000B', '99000X', '99000Y', '99371A', '99402A', '99402B', '99402C',
             '99484B', '99556A', '99679H', '99679S', '99691E', '99691H']

    while True:

        while number < len(list1):

            teamloop = list1[number]
            # ['sheet_%d' % sheetnb].write = book.add_sheet(teamloop, cell_overwrite_ok= True)
            print('')
            print(teamloop)
            print('')
            number += 1
            sheetline += 1

            r = urlopen(VEXDB_API_RANK + str(teamloop) + VEX_SEASON)
            text = r.read()
            json_dict = json.loads(text)
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
                    sheet2.write(sheetline, 6, "Positive", STYLE_1)
                elif int(datawins) < int(datalosses):
                    sheet2.write(sheetline, 6, "Negative", STYLE_2)
                output.append(line)

                # ['sheet' + str(number)].write(1, 1,teamloop)
                # ['sheet' + str(number)].write(2, 1,line )

                sheet2.write(sheetline, 0, datateam)
                sheet2.write(sheetline, 1, datawins)
                sheet2.write(sheetline, 2, datalosses)
                sheet2.write(sheetline, 3, dataap)
                sheet2.write(sheetline, 4, datarank)
                sheet2.write(sheetline, 5, datamaxscore)
                sheetline += 1

                # print(line)

                time.sleep(sleep_timer)

            # pprint.pprint(output)
            book.save("Data.xls")

            print('')

            decimal = (time.time() - start)
            decimal = Decimal.from_float(decimal).quantize(Decimal('0.0'))

            ave = (float(decimal) / (int(number)))
            ave = Decimal.from_float(ave).quantize(Decimal('0.0'))

            eta = (float(ave) * (int(len(list1) - (int(number)))))
            etatomin = (float(eta) / 60)
            etatomin = Decimal.from_float(etatomin).quantize(Decimal('0.0'))

            print(str(number) + "/" + str(len(list1)) + " Finished, Used " + str(decimal) + " seconds. Average " + str(
                ave) + " seconds each. ETA: " + str(etatomin) + " mins.")

            print()
            print('wait ' + str(sleep_timer) +
                  ' sec for next request to server...')

        if number >= 5:
            number = 0
            sheetline = 1
            print('')
            print('reset and xls saved!')
            time.sleep(2)

    # time.sleep(1)


def excel_get_all_data():  # 203
    time.sleep(1)
    number = 0
    sheetline = 0
    start = time.time()

    # TODO(Yifei):find a way to get this list from Vex.io
    list1 = ['2S', '2U', '5S', '10N', '12C', '12E', '12F', '12G', '12J', '39A', '39J', '39K', '39W', '39Y', '46B',
             '56C', '60X', '62A', '66A', '81K', '81Y', '91C', '109A', '114T', '127X', '134C', '134D', '134E', '134G',
             '136N', '162A', '169A', '169C', '169E', '169Y', '170A', '177V', '180A', '180X', '183Z', '185A', '202Z',
             '231X', '244A', '244B', '278E', '285X', '288A', '306X', '315G', '315J', '315X', '315Z', '321A', '323G',
             '323Y', '333R', '343X', '355D', '355E', '355M', '356A', '356B', '356C', '359A', '359X', '363A', '365X',
             '398A', '409A', '464M', '507D', '523X', '536C', '536E', '546A', '574C', '574F', '590G', '590V', '598B',
             '599A', '621A', '624H', '624K', '643T', '643Z', '660A', '666U', '666X', '675D', '686A', '709S', '815J',
             '817A', '824Z', '901C', '917F', '920C', '929H', '929U', '929X', '934Z', '970A', '986A', '1008M', '1010B',
             '1010N', '1010X', '1028A', '1028B', '1028Z', '1045A', '1045B', '1064A', '1069E', '1104S', '1115A', '1138B',
             '1193Z', '1200Z', '1233F', '1235C', '1248B', '1264D', '1267C', '1275A', '1275B', '1275D', '1320B', '1320C',
             '1320D', '1344A', '1353C', '1356B', '1410A', '1429B', '1437Z', '1460A', '1483B', '1492W', '1492X', '1492Y',
             '1492Z', '1505R', '1533M', '1575X', '1588A', '1588D', '1617A', '1690X', '1718A', '1727E', '1727F', '1784Z',
             '1814D', '1859S', '1859W', '1859X', '1961K', '1961N', '1961U', '1961X', '1965A', '1965R', '1965T', '1970K',
             '1973B', '2011C', '2011E', '2011F', '2019B', '2019F', '2030A', '2030B', '2114X', '2114Z', '2131E', '2131M',
             '2131R', '2131W', '2131X', '2142D', '2223Z', '2235A', '2250K', '2263C', '2284A', '2284B', '2297A', '2316A',
             '2360S', '2360V', '2373A', '2396A', '2435B', '2435W', '2442B', '2442C', '2456A', '2496V', '2496Y', '2560E',
             '2560S', '2567X', '2612A', '2616D', '2616J', '2616Y', '2719B', '2719D', '2775B', '2777V', '2777W', '2886B',
             '2900B', '2900C', '2900G', '2921S', '2941A', '2979A', '2990B', '2993M', '3018A', '3018V', '3050A', '3050C',
             '3118B', '3141S', '3159A', '3260S', '3264N', '3269A', '3269B', '3273B', '3314K', '3348B', '3388X', '3389D',
             '3547A', '3553C', '3631C', '3631Z', '3682A', '3701A', '3767A', '3767X', '3796C', '3815M', '3818D', '3946E',
             '3946W', '4004X', '4057C', '4104A', '4104C', '4142A', '4147A', '4148A', '4148D', '4154A', '4154B', '4169J',
             '4253J', '4255A', '4305A', '4306A', '4318B', '4364H', '4403A', '4409A', '4410C', '4411Y', '4454A', '4478V',
             '4478X', '4549B', '4610C', '4610Z', '4621B', '4805F', '4815A', '4815B', '4828B', '4911A', '5062A', '5090Z',
             '5106C', '5139A', '5139D', '5221T', '5225A', '5245A', '5300A', '5327A', '5327B', '5327C', '5327X', '5408A',
             '5588D', '5588E', '5588R', '5691Y', '5735B', '5735K', '5776A', '5776E', '5776T', '5864B', '5956B', '5956G',
             '5999W', '6007X', '6008A', '6008D', '6008Z', '6106B', '6106C', '6109C', '6121C', '6135H', '6135W', '6210Z',
             '6277B', '6299B', '6358A', '6358C', '6403A', '6403B', '6603V', '6627A', '6627B', '6627D', '6671X', '6715B',
             '6724B', '6741A', '6822B', '6842A', '6916C', '6916E', '6916H', '6978J', '6980A', '7110A', '7110Y', '7110Z',
             '7121D', '7121E', '7209X', '7221R', '7221T', '7232X', '7258A', '7258B', '7316D', '7316F', '7323A', '7368W',
             '7386X', '7432B', '7432C', '7447B', '7458F', '7479A', '7536B', '7546A', '7618B', '7618C', '7682E', '7682S',
             '7700R', '7700S', '7776B', '7842B', '7842F', '7853A', '7856A', '7862B', '7862X', '7884B', '7884D', '7884E',
             '7975F', '7983V', '7984B', '8000A', '8000B', '8000C', '8000E', '8044A', '8059A', '8059D', '8059J', '8059X',
             '8059Z', '8079A', '8086A', '8110B', '8110C', '8110X', '8110Z', '8114A', '8176B', '8192B', '8192C', '8192D',
             '8223A', '8232X', '8261A', '8330A', '8331A', '8349E', '8373H', '8387B', '8387C', '8447A', '8451D', '8452A',
             '8568A', '8659G', '8669A', '8675A', '8691E', '8787A', '8800X', '8825S', '8855B', '8861C', '8931F', '8995M',
             '9020A', '9031H', '9060A', '9060B', '9060C', '9080R', '9090A', '9144E', '9185B', '9225C', '9228A', '9343C',
             '9364A', '9364C', '9364D', '9364E', '9409A', '9409B', '9409C', '9421X', '9457B', '9545W', '9551B', '9553E',
             '9594C', '9605A', '9623E', '9727A', '9784B', '9823C', '9873B', '9922A', '9922Z', '9932B', '9932E', '9932F',
             '9932G', '9973A', '10173W', '10300X', '10955M', '11124R', '11495A', '12298A', '15376A', '16101Z', '17071B',
             '17090A', '18554B', '20610A', '20785A', '20785B', '21246X', '21508A', '22699A', '23880B', '25461Z',
             '26982E', '27183R', '34000E', '34000M', '34203X', '34760A', '35211C', '35960A', '41364A', '41998A',
             '43775B', '44244C', '48180S', '48327M', '48667A', '48778A', '49181C', '49181N', '49181T', '49181U',
             '49450A', '50505A', '51140A', '51140B', '51581A', '53999E', '55563E', '57418A', '58072A', '59990Z',
             '61499U', '61499Y', '61601A', '62019P', '62440F', '62993A', '64000B', '64040B', '64846A', '67292A',
             '68211A', '68555A', '71303A', '72477A', '76209G', '76565A', '77177B', '77321J', '77788J', '81118P',
             '81785K', '86000A', '86868R', '87217R', '91709A', '93063A', '96666D', '96671B', '96671X', '97038A',
             '97140A', '97301C', '97371A', '97871A', '97934U', '98177B', '98177C', '98548B', '98725B', '98744A',
             '98807A', '99000A', '99000B', '99000X', '99000Y', '99371A', '99402A', '99402B', '99402C', '99484B',
             '99679H', '99679S']

    while True:

        while number < int(len(list1)):

            teamloop = list1[number]
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

            r = urlopen(VEXDB_API_RANK + teamloop + VEX_SEASON)
            text = r.read()

            json_dict = json.loads(text)

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

            r = urlopen(VEXDB_API_MATCHES + teamloop + VEX_SEASON)
            text = r.read()
            # pprint.pprint(json.loads(text))
            json_dict = json.loads(text)
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

                # pprint.pprint(output)

                time.sleep(sleep_timer)

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

            eta = (float(ave) * (int(len(list1) - (int(number)))))
            etatomin = (float(eta) / 60)
            etatomin = Decimal.from_float(etatomin).quantize(Decimal('0.0'))

            print(str(number) + "/" + str(len(list1)) + " Finished, Used " + str(decimal) + " seconds. Average " + str(
                ave) + " seconds each. ETA: " + str(etatomin) + " mins.")
            print()
            book.save("Data" + ".xls")

        if number >= 5:
            number = 0
            sheetline = 1
            print('')
            print('reset and xls saved!')


def excel_get_all_bugs():  # 204
    time.sleep(1)
    number = 0
    sheetline = 0
    start = time.time()
    list1 = ['7386A', '8000X', '8000Z', '19771B', '30638A', '36632A', '37073A', '60900A', '76921B', '99556A', '99691E',
             '99691H']

    while True:
        while number < int(len(list1)):
            teamloop = list1[number]
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

            r = urlopen(VEXDB_API_RANK + teamloop + VEX_SEASON)
            text = r.read()

            json_dict = json.loads(text)

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

            # pprint.pprint(output)

            r = urlopen(VEXDB_API_MATCHES + teamloop + VEX_SEASON)
            text = r.read()
            # pprint.pprint(json.loads(text))
            json_dict = json.loads(text)
            # print('\n')
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

                # sheetline += 1

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

                # pprint.pprint(output)

                time.sleep(sleep_timer)

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

            eta = float(ave) * (int(len(list1) - (int(number))))
            etatomin = (float(eta) / 60)
            etatomin = Decimal.from_float(etatomin).quantize(Decimal('0.0'))

            print(str(number) + "/" + str(len(list1)) + " Finished, Used " + str(decimal) + " seconds. Average " + str(
                ave) + " seconds each. ETA: " + str(etatomin) + " mins.")
            print()
            book.save("Data" + ".xls")

        if number >= 5:
            number = 0
            sheetline = 1
            print('')
            print('reset and xls saved!')


def excel_get_we_need():  # 205
    time.sleep(1)
    number = 0
    sheetline = 0
    start = time.time()
    list1 = ['2U', '81K', '169A', '365X', '624K', '934Z', '1064A', '1437Z', '1961U', '2360S', '2719B', '3269B', '3767A',
             '4815B', '5139A', '6627A', '6741A', '7258B', '7536B', '7853A', '8110B', '8192B', '9060C', '9228A', '9551B',
             '9932E', '10955M', '17071B', '35211C', '97934U', '98807A']

    while True:  # Todo(Yifei): What is this loop for?

        while number < int(len(list1)):  # TODO(Yifei): Use for loop instead

            teamloop = list1[number]
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

            r = urlopen(VEXDB_API_RANK + teamloop + VEX_SEASON)
            text = r.read()

            json_dict = json.loads(text)

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

            # pprint.pprint(output)

            r = urlopen(VEXDB_API_MATCHES + teamloop + VEX_SEASON)  # TODO(Yifei): Turn this into a function

            text = r.read()

            # pprint.pprint(json.loads(text))

            json_dict = json.loads(text)

            # print('\n')

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

                # pprint.pprint(output)

                time.sleep(sleep_timer)

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

            eta = float(ave) * (int(len(list1) - (int(number))))
            etatomin = (float(eta) / 60)
            etatomin = Decimal.from_float(etatomin).quantize(Decimal('0.0'))

            print(str(number) + "/" + str(len(list1)) + " Finished, Used " + str(decimal) + " seconds. Average " + str(
                ave) + " seconds each. ETA: " + str(etatomin) + " mins.")
            print()
            book.save("Data" + ".xls")

        if number >= 5:
            number = 0
            sheetline = 1
            print('')
            print('reset and xls saved!')


def excel_scan_world():
    time.sleep(1)
    number = 0
    sheetline = 0
    start = time.time()
    list1 = ['2S', '2U', '5S', '10N', '12C', '12E', '12F', '12G', '12J', '39A', '39J', '39K', '39W', '39Y', '46B',
             '56C', '60X', '62A', '66A', '81K', '81Y', '91C', '109A', '114T', '127X', '134C', '134D', '134E', '134G',
             '136N', '162A', '169A', '169C', '169E', '169Y', '170A', '177V', '180A', '180X', '183Z', '185A', '202Z',
             '231X', '244A', '244B', '278E', '285X', '288A', '306X', '315G', '315J', '315X', '315Z', '321A', '323G',
             '323Y', '333R', '343X', '355D', '355E', '355M', '356A', '356B', '356C', '359A', '359X', '363A', '365X',
             '398A', '409A', '464M', '507D', '523X', '536C', '536E', '546A', '574C', '574F', '590G', '590V', '598B',
             '599A', '621A', '624H', '624K', '643T', '643Z', '660A', '666U', '666X', '675D', '686A', '709S', '815J',
             '817A', '824Z', '901C', '917F', '920C', '929H', '929U', '929X', '934Z', '970A', '986A', '1008M', '1010B',
             '1010N', '1010X', '1028A', '1028B', '1028Z', '1045A', '1045B', '1064A', '1069E', '1104S', '1115A', '1119S',
             '1138B', '1193Z', '1200Z', '1233F', '1235C', '1248B', '1264D', '1267C', '1275A', '1275B', '1275D', '1320B',
             '1320C', '1320D', '1344A', '1353C', '1356B', '1410A', '1429B', '1437Z', '1460A', '1483B', '1492W', '1492X',
             '1492Y', '1492Z', '1505R', '1533M', '1575X', '1588A', '1588D', '1617A', '1690X', '1718A', '1727E', '1727F',
             '1784Z', '1814D', '1859S', '1859W', '1859X', '1961K', '1961N', '1961U', '1961X', '1965A', '1965R', '1965T',
             '1970K', '1973B', '2011C', '2011E', '2011F', '2019B', '2019F', '2030A', '2030B', '2114X', '2114Z', '2131E',
             '2131M', '2131R', '2131W', '2131X', '2142D', '2223Z', '2235A', '2250K', '2263C', '2284A', '2284B', '2297A',
             '2316A', '2360S', '2360V', '2373A', '2396A', '2435B', '2435W', '2442B', '2442C', '2456A', '2496V', '2496Y',
             '2560E', '2560S', '2567X', '2612A', '2616D', '2616J', '2616Y', '2719B', '2719D', '2775B', '2777V', '2777W',
             '2886B', '2900B', '2900C', '2900G', '2921S', '2941A', '2979A', '2990B', '2993M', '3018A', '3018V', '3050A',
             '3050C', '3118B', '3141S', '3159A', '3260S', '3264N', '3269A', '3269B', '3273B', '3314K', '3348B', '3388X',
             '3389D', '3547A', '3553C', '3631C', '3631Z', '3682A', '3701A', '3767A', '3767X', '3796C', '3815M', '3818D',
             '3946E', '3946W', '4004X', '4057C', '4104A', '4104C', '4142A', '4147A', '4148A', '4148D', '4154A', '4154B',
             '4169J', '4253J', '4255A', '4305A', '4306A', '4318B', '4364H', '4403A', '4409A', '4410C', '4411Y', '4454A',
             '4478V', '4478X', '4549B', '4610C', '4610Z', '4621B', '4805F', '4815A', '4815B', '4828B', '4911A', '5062A',
             '5090Z', '5106C', '5139A', '5139D', '5221T', '5225A', '5245A', '5300A', '5327A', '5327B', '5327C', '5327X',
             '5408A', '5588D', '5588E', '5588R', '5691Y', '5735B', '5735K', '5776A', '5776E', '5776T', '5864B', '5956B',
             '5956G', '5999W', '6007X', '6008A', '6008D', '6008Z', '6106B', '6106C', '6109C', '6121C', '6135H', '6135W',
             '6210Z', '6277B', '6299B', '6358A', '6358C', '6403A', '6403B', '6603V', '6627A', '6627B', '6627D', '6671X',
             '6715B', '6724B', '6741A', '6822B', '6842A', '6916C', '6916E', '6916H', '6978J', '6980A', '7110A', '7110Y',
             '7110Z', '7121D', '7121E', '7209X', '7221R', '7221T', '7232X', '7258A', '7258B', '7316D', '7316F', '7323A',
             '7368W', '7386A', '7386X', '7432B', '7432C', '7447B', '7458F', '7479A', '7536B', '7546A', '7618B', '7618C',
             '7682E', '7682S', '7686A', '7700R', '7700S', '7776B', '7842B', '7842F', '7853A', '7856A', '7862B', '7862X',
             '7884B', '7884D', '7884E', '7975F', '7983V', '7984B', '8000A', '8000B', '8000C', '8000E', '8000X', '8000Z',
             '8044A', '8059A', '8059D', '8059J', '8059X', '8059Z', '8079A', '8086A', '8110B', '8110C', '8110X', '8110Z',
             '8114A', '8176B', '8192B', '8192C', '8192D', '8223A', '8232X', '8261A', '8330A', '8331A', '8349E', '8373H',
             '8387B', '8387C', '8447A', '8451D', '8452A', '8568A', '8659G', '8669A', '8675A', '8691E', '8787A', '8800X',
             '8825S', '8855B', '8861C', '8931F', '8995M', '9020A', '9031H', '9060A', '9060B', '9060C', '9080R', '9090A',
             '9144E', '9185B', '9225C', '9228A', '9343C', '9364A', '9364C', '9364D', '9364E', '9409A', '9409B', '9409C',
             '9421X', '9457B', '9545W', '9551B', '9553E', '9594C', '9605A', '9623E', '9727A', '9784B', '9823C', '9873B',
             '9922A', '9922Z', '9932B', '9932E', '9932F', '9932G', '9973A', '10173W', '10300X', '10955M', '11124R',
             '11495A', '12298A', '15376A', '16101Z', '17071B', '17090A', '18554B', '19771B', '20610A', '20785A',
             '20785B', '21246X', '21508A', '22699A', '23880B', '25461Z', '26982E', '27183R', '30638A', '34000E',
             '34000M', '34203X', '34760A', '35211C', '35960A', '36632A', '37073A', '41364A', '41998A', '43775B',
             '44244C', '48180S', '48327M', '48667A', '48778A', '49181C', '49181N', '49181T', '49181U', '49450A',
             '50505A', '51140A', '51140B', '51581A', '53999E', '55563E', '57418A', '58072A', '59990Z', '60900A',
             '61499U', '61499Y', '61601A', '62019P', '62440F', '62993A', '64000B', '64040B', '64846A', '67292A',
             '68211A', '68555A', '71303A', '72477A', '76209G', '76565A', '76921B', '77177B', '77321J', '77788J',
             '81118P', '81785K', '86000A', '86868R', '87217R', '91709A', '93063A', '96666D', '96671B', '96671X',
             '97038A', '97140A', '97301C', '97371A', '97871A', '97934U', '98177B', '98177C', '98548B', '98725B',
             '98744A', '98807A', '99000A', '99000B', '99000X', '99000Y', '99371A', '99402A', '99402B', '99402C',
             '99484B', '99556A', '99679H', '99679S', '99691E', '99691H']

    while True:

        while number < int(len(list1)):

            teamloop = list1[number]
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

            r = urlopen(VEXDB_API_RANK + teamloop + VEX_SEASON + '&sku=RE-VRC-17-3805')
            text = r.read()

            json_dict = json.loads(text)

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

            # pprint.pprint(output)
            r = urlopen(VEXDB_API_MATCHES + teamloop + VEX_SEASON)
            text = r.read()
            # pprint.pprint(json.loads(text))
            json_dict = json.loads(text)
            # print('\n')
            output = []
            loop = -10000
            # 1-10000 For testing, should be 0

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

                # pprint.pprint(output)

                time.sleep(0.1)

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

            eta = float(ave) * (int(len(list1) - (int(number))))
            etatomin = (float(eta) / 60)
            etatomin = Decimal.from_float(etatomin).quantize(Decimal('0.0'))

            print(str(number) + "/" + str(len(list1)) + " Finished, Used " + str(decimal) + " seconds. Average " + str(
                ave) + " seconds each. ETA: " + str(etatomin) + " mins.")
            print()
            book.save("Data" + ".xls")

        if number >= 5:
            number = 0
            sheetline = 1
            print('\n reset and xls saved!')


# Need to test when competition start


def excel_team_matches():
    name = input('Team #?\n')
    print('Checking, TEAM %s.' % name)

    r = urlopen(VEXDB_API_MATCHES + name + VEX_SEASON)
    text = r.read()
    pprint.pprint(json.loads(text))
    json_dict = json.loads(text)
    print('\n')
    output = []
    for r in json_dict["result"]:
        line = '{}: Match{} Round{} || Red Alliance 1 = {} Red Alliance 2 = {} Red Alliance 3 = {} Red Sit = {} || ' \
               'Blue Alliance 1 = {} Blue Alliance 2 = {} Blue Alliance 3 = {} Blue Sit = {} || Red Score = {} Blue ' \
               'Score = {}' \
            .format(r["sku"], r["matchnum"], r["round"], r["red1"], r["red2"], r["red3"], r["redsit"], r["blue1"],
                    r["blue2"], r["blue3"], r["bluesit"], r["redscore"], r["bluescore"])
        output.append(line)
    pprint.pprint(output)
    time.sleep(1)
    return None


def search_team_current_season():
    name = str(input('Team #?\n'))
    print('Checking, TEAM %s.' % name)
    r = urlopen(VEXDB_API_RANK + name + VEX_SEASON)
    text = r.read()
    pprint.pprint(json.loads(text))
    json_dict = json.loads(text)
    print('\n')
    output = []
    for r in json_dict["result"]:
        line = "Team = {} Wins = {} Losses = {} AP = {} Ranking in Current Match = {} Highest Score = {}" \
            .format(r["team"], r["wins"], r["losses"], r["ap"], r["rank"], r["max_score"])
        output.append(line)
    pprint.pprint(output)
    time.sleep(1)
    return None


def get_all_data():
    # getalldata
    print("This will show the recent three matches.")
    name = str(input('Team #?\n'))
    print('Checking, TEAM %s.' % name)
    r = urlopen(VEXDB_API_RANK + name + VEX_SEASON)
    text = r.read()
    # pprint.pprint(json.loads(text))
    json_dict = json.loads(text)
    # print('\n')
    output = []
    for r in json_dict["result"]:
        line = "Team = {} Wins = {} Losses = {} AP = {} Ranking in Current Match = {} Highest Score = {}" \
            .format(r["team"], r["wins"], r["losses"], r["ap"], r["rank"], r["max_score"])
        output.append(line)
    pprint.pprint(output)
    r = urlopen(VEXDB_API_MATCHES + name + VEX_SEASON)
    text = r.read()
    # pprint.pprint(json.loads(text))
    json_dict = json.loads(text)
    # print('\n')
    output = []
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
        output.append(line)
        pprint.pprint(output)
        # book.save(name + ".xls")
    # time.sleep(1)
    return None


def time_is_out():
    # Input Team
    GlobalVar.inputmode = str(
        input("Type in the preset value or 6 teams separate by ,\n"))

    print(
        "TR1: " + GlobalVar.teamr1 + " TR2: " + GlobalVar.teamr2 + " TR3: " + GlobalVar.teamr3 + " || TB1: "
        + GlobalVar.teamb1 + " TB2: " + GlobalVar.teamb2 + " TB3: " + GlobalVar.teamb3)

    if str(GlobalVar.teamr1) != "":
        GlobalVar.teamsent = GlobalVar.teamr1
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
        GlobalVar.teamsent = GlobalVar.teamr2
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
        GlobalVar.teamsent = GlobalVar.teamr3
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
        GlobalVar.teamsent = GlobalVar.teamb1
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
        GlobalVar.teamsent = GlobalVar.teamb2
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
        GlobalVar.teamsent = GlobalVar.teamb3
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

    # print("Skill is average of all this season. Auto is the previous competition. (Should be the state final)")
    # print("Ranking is (10-Ranking), if the team is not the first 10th, it will show as 0.")

    graphbubble()  # pass value use arg instead of global

    return None


def team_skill():
    r = urlopen('https://api.vexdb.io/v1/get_skills?team=' + GlobalVar.teamsent + VEX_SEASON)
    text = r.read()
    json_dict = json.loads(text)
    # output = []
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


def team_sent():
    count = 0
    GlobalVar.winsave = 0
    r = urlopen(VEXDB_API_RANK + GlobalVar.teamsent + VEX_SEASON)
    text = r.read()
    json_dict = json.loads(text)
    for r in json_dict["result"]:
        # line = '{}'.format(r["wins"])
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


def team_current():  # can be part of teamsent()
    GlobalVar.currentranking = 0
    GlobalVar.currentwins = 0
    GlobalVar.currentlosses = 0
    r = urlopen(VEXDB_API_RANK + GlobalVar.teamsent + VEX_SEASON + GlobalVar.CONST_match)
    text = r.read()
    json_dict = json.loads(text)
    for r in json_dict["result"]:
        # line = '{}'.format(r["rank"], r["wins"], r["losses"])
        # output.append(line)
        GlobalVar.currentranking = '{}'.format(r["rank"])
        GlobalVar.currentwins = '{}'.format(r["wins"])
        GlobalVar.currentlosses = '{}'.format(r["losses"])
    teamap()


def teamap():
    aptotal = 0
    count = 0
    r = urlopen(VEXDB_API_RANK + GlobalVar.teamsent + VEX_SEASON)
    text = r.read()
    json_dict = json.loads(text)

    for r in json_dict["result"]:
        # line = '{}'.format(r["ap"])
        # output.append(line)
        teammap = '{}'.format(r["ap"])
        count += 1

        if int(teammap) > 25:
            diff = (int(teammap) - 25) * 0.2
            teammap = 25 + float(diff)
            print("Balance over 25, " + str(diff))
        aptotal = int(aptotal) + int(teammap)
        GlobalVar.apave = int(aptotal) / int(count)

        if teammap == "" or teammap == "":
            print("break cuz blank")
            count -= 1
            teamranking()
    teamranking()


def teamranking():
    GlobalVar.rankave = 0
    count = 0
    r = urlopen(VEXDB_API_RANK + GlobalVar.teamsent + VEX_SEASON)
    text = r.read()
    json_dict = json.loads(text)
    for r in json_dict["result"]:
        # line = '{}'.format(r["rank"])
        # output.append(line)
        team_ranking = '{}'.format(r["rank"])
        count += 1
        ranktotal = int(
            ranktotal) + int(team_ranking)
        GlobalVar.rankave = float(ranktotal) / int(count)

        if team_ranking == "" or team_ranking == "":
            print("break cuz blank")
            count -= 1
            GlobalVar.rankave = float(ranktotal) / int(count)
            team_highest()
        GlobalVar.rankave = float(team_ranking) / int(count)
    team_highest()


def team_highest():
    highesttotal = 0
    GlobalVar.highestave = 0
    count = 0
    r = urlopen(VEXDB_API_RANK + GlobalVar.teamsent + VEX_SEASON)
    text = r.read()
    json_dict = json.loads(text)
    for r in json_dict["result"]:
        # line = '{}'.format(r["max_score"])
        # output.append(line)
        team_highest = '{}'.format(r["max_score"])
        count += 1
        highesttotal = int(
            highesttotal) + int(team_highest)
        GlobalVar.highestave = int(highesttotal) / count
        if team_highest == "":
            print("break cuz blank")
            count -= 1
            GlobalVar.highestave = float(highesttotal) / int(count)
            teampr()
        GlobalVar.highestave = float(highesttotal) / int(count)
    teampr()


def teampr():
    GlobalVar.oprtotal = 0
    dprtotal = 0
    r = urlopen(VEXDB_API_RANK + GlobalVar.teamsent + VEX_SEASON)
    text = r.read()
    json_dict = json.loads(text)
    count = 0
    for r in json_dict["result"]:
        # line = '{} {}'.format(r["opr"], r["dpr"])
        # output.append(line)
        teamopr = '{}'.format(r["opr"])
        teamdpr = '{}'.format(r["dpr"])
        teamopr = (float(teamopr) / 5)
        teamdpr = (float(teamdpr) / 5)
        count += 1
        GlobalVar.oprtotal = float(
            GlobalVar.oprtotal) + float(teamopr)
        GlobalVar.oprave = float(GlobalVar.oprtotal) / int(count)
        dprtotal = float(
            dprtotal) + float(teamdpr)
        GlobalVar.dprave = float(dprtotal) / int(count)

        if teamdpr == "" or teamopr == "":
            print("break cuz blank")
            count -= 1
            teamccwm()

        teamccwm()

        '''
        for r in json_dict["result"]:
            line = '{} {}'.format(r["opr"], r["dpr"])
            # output.append(line)
            teamopr = '{}'.format(r["opr"])
            teamdpr = '{}'.format(r["dpr"])
            teamopr = (float(teamopr) / 5)
            teamdpr = (float(teamdpr) / 5)
        '''


def teamccwm():
    ccwmtotal = 0
    GlobalVar.ccwmave = 0
    r = urlopen(VEXDB_API_RANK + GlobalVar.teamsent + VEX_SEASON)
    text = r.read()
    json_dict = json.loads(text)
    count = 0
    for r in json_dict["result"]:
        # line = '{}'.format(r["ccwm"])
        # output.append(line)
        teamccwm = '{}'.format(r["ccwm"])
        count += 1
        ccwmtotal = float(
            ccwmtotal) + float(teamccwm)
        GlobalVar.ccwmave = float(ccwmtotal) / int(count)
        if teamccwm == "" or teamccwm == "":
            print("break cuz blank")
            count -= 18
            break


def graphbubble():  # it should be part of "timeisout"
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
    GlobalVar.teamr1highest = round(
        float(int(GlobalVar.teamr1highest) / 17), 1)
    GlobalVar.teamr2highest = round(
        float(int(GlobalVar.teamr2highest) / 17), 1)
    GlobalVar.teamr3highest = round(
        float(int(GlobalVar.teamr3highest) / 17), 1)

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
    try:
        os.remove("graph/" + GlobalVar.inputmode + ".png")
        print("Previous deleted.")
        time.sleep(1)
    except OSError:
        print("something is not right")
        pass
    plt.xlabel(
        "Skill / [Defensive]")
    plt.ylabel("AP / [Offensive]")
    plt.title(
        "Red: " + GlobalVar.teamr1 + " " + GlobalVar.teamr2 + " " + GlobalVar.teamr3 +
        " Blue: " + GlobalVar.teamb1 + " " +
        GlobalVar.teamb2 + " " + GlobalVar.teamb3,
        loc="left")
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
    plt.savefig("graph/" + GlobalVar.inputmode + ".png")
    print("Graph poped and saved.")
    plt.show()


def answer():
    # answerr = 0
    # answerb = 0

    teamrexist = 0
    teambexist = 0

    # teamrskill = float(GlobalVar.teamr1skillout) + float(GlobalVar.teamr2skillout) + float(
    #     GlobalVar.teamr3skillout)
    # teambskill = float(GlobalVar.teamb1skillout) + float(GlobalVar.teamb2skillout) + float(
    #     GlobalVar.teamb3skillout)
    # teamrave = (float(GlobalVar.teamrskill) / 3)
    # teambave = (float(GlobalVar.teambskill)) / 3

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

    time.sleep(2)
    input("Press Any Key to Continue\n")


# Start!

while True:
    mode = int(input(
        "Mode \n 1.Scan Team Matches \n 2.Excel Functions [Not Finished] \n 3.Search Team Season History  "
        "\n 8.Get Important Info For a Team \n 9.Change Log\n 0.Quit \n"))
    if mode == 1:
        print("Mode = Scan Team Matches")
        time.sleep(0.3)
        scan_team_matches()
    elif mode == 2:
        print("Mode = Excels")
        # sleep_timer = float(input("Set Sleep Time\n"))
        print(
            "1.Scan Teams \n2.Scan Matches [Don't use this]\n3.Write Team Important Data\n4.Don't Ues This\n5.Can "
            "Specific Match [PreSet World Championship]\n6.Get We Need")
        time.sleep(0.3)
        excelmode = int(input())
        if excelmode == 1:
            print("Mode = Scan Teams and Write to Excel")
            time.sleep(0.3)
            excel_scan_teams()
        elif excelmode == 2:
            print("Mode = Write Team Matches [Don't use this]")
            time.sleep(0.3)
            excel_team_matches()
        elif excelmode == 3:
            print("Mode = Write Team Important Data in Excel")
            time.sleep(0.3)
            excel_get_all_data()
        elif excelmode == 4:
            print("Mode = Scan Bugged Team [It will crash]")
            time.sleep(0.3)
            excel_get_all_bugs()
        elif excelmode == 5:
            print("Mode = Scan World Championship")
            time.sleep(0.3)
            excel_scan_world()
        elif excelmode == 6:
            print("Mode = Scan We Need")
            time.sleep(0.3)
            excel_get_we_need()
    elif mode == 3:
        print("Mode = Search Team History : Current Season")
        time.sleep(0.3)
        search_team_current_season()
    elif mode == 4:
        print("Bubble!")
        time_is_out()
        answer()
    elif mode == 8:
        print("Mode = Get Important Data")
        time.sleep(0.3)
        get_all_data()
    elif mode == 0:
        print("Thanks for using it!")
        time.sleep(0.3)
        quit()
    else:
        print("Mode Unknown")
        time.sleep(1)
