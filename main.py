import json
import os
import pprint
import time
import webbrowser
import xlwt
import matplotlib.pyplot as plt
import pandas as pd
from decimal import getcontext, Decimal
from math import pi

# import errno
# import numpy as np
# import seaborn as sns


# preload
getcontext().prec = 6
sleeptimer = 0
book = xlwt.Workbook(encoding="utf-8")

#title = ["#Cover", "#Matches", "#Important Data", "#Blank", "#For World", "#What We Need", "#Team Spot 1"
#    , "#Team Spot 2", "#Team Spot 3", "#Team Spot 4", "#Bugged Teams"]


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
timenow = "Last Update:" + time.strftime("%c")
sheet1.write(2, 1, timenow)
sheet1.write(3, 1,
             "Because of there are no data for these teams: 1119S, 7386A, 8000X, 8000Z, 19771B, 30638A, 36632A, "
             "37073A, 60900A, 76921B, 99556A, 99691E, 99691H are not include in the sheet #Important Data")
style1 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour red;''font: colour white, bold True;')
style2 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour blue;''font: colour white, bold True;')
style3 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour pink;''font: colour white, bold True;')
style4 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour pale_blue;''font: colour white, bold True;')
stylered = xlwt.easyxf(
    'font: colour red, bold True;')
styleblue = xlwt.easyxf(
    'font: colour blue, bold True;')
styleblank = xlwt.easyxf(
    'pattern: pattern solid, fore_colour black;''font: colour white, bold True;')
styleb = xlwt.easyxf(
    'font: colour black, bold True;')
style70 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour red;''font: colour white, bold True;')
style50 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour light_orange;''font: colour white, bold True;')
style30 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour pale_blue;''font: colour white, bold True;')
style0 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour bright_green;''font: colour black, bold True;')

sheet2.write(0, 0, "Team")
sheet2.write(0, 1, "Wins")
sheet2.write(0, 2, "Losses")
sheet2.write(0, 3, "AP")
sheet2.write(0, 4, "Ranking")
sheet2.write(0, 5, "Highest")
sheet2.write(0, 6, "Result")


class global_var:

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
    winstotal = 0
    apave = 0
    aptotal = 0
    oprave = 0
    oprtotal = 0
    dprave = 0
    dprtotal = 0
    rankave = 0
    ranktotal = 0
    highestave = 0
    ccwmave = 0


def scanteammatches():
    name = input('Team #?\n')
    print('Checking, TEAM %s.' % name)

    from urllib.request import urlopen

    r = urlopen('https://api.vexdb.io/v1/get_matches?team=' +
                name + '&season=Turning%20Point')

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


def excelscanteams():  # 201
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

            from urllib.request import urlopen

            r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                        str(teamloop) + '&season=Turning%20Point')

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
                    sheet2.write(sheetline, 6, "Positive", style1)
                elif int(datawins) < int(datalosses):
                    sheet2.write(sheetline, 6, "Negative", style2)
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

                time.sleep(sleeptimer)

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
            print('wait ' + str(sleeptimer) +
                  ' sec for next request to server...')

        if number >= 5:
            number = 0
            sheetline = 1
            print('')
            print('reset and xls saved!')
            time.sleep(2)

    # time.sleep(1)


def excelgetalldata():  # 203
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

            from urllib.request import urlopen
            r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                        teamloop + '&season=Turning%20Point')
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
                sheet3.write(sheetline, 6, "Positive", style1)
            elif int(datawins) < int(datalosses):
                sheet3.write(sheetline, 6, "Negative", style2)

            sheetline += 1

            # pprint.pprint(output)

            from urllib.request import urlopen

            r = urlopen('https://api.vexdb.io/v1/get_matches?team=' +
                        teamloop + '&season=Turning%20Point')

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
                sheet3.write(sheetline, 2, datared1, stylered)
                sheet3.write(sheetline, 3, datared2, stylered)
                sheet3.write(sheetline, 4, datared3, stylered)
                sheet3.write(sheetline, 5, dataredsit, stylered)
                sheet3.write(sheetline, 6, datablue1, styleblue)
                sheet3.write(sheetline, 7, datablue2, styleblue)
                sheet3.write(sheetline, 8, datablue3, styleblue)
                sheet3.write(sheetline, 9, databluesit, styleblue)
                sheet3.write(sheetline, 10, dataredsc, stylered)
                sheet3.write(sheetline, 11, databluesc, styleblue)
                sheet3.write(sheetline, 12, datateam + " =", styleb)

                if int(dataredsc) > int(databluesc):
                    sheet3.write(sheetline, 14, "Red", style1)
                elif int(dataredsc) < int(databluesc):
                    sheet3.write(sheetline, 14, "Blue", style2)

                if int(dataredsc) + 20 < int(databluesc):
                    sheet3.write(sheetline, 14, "Blue Easy", style4)
                elif int(dataredsc) - 20 > int(databluesc):
                    sheet3.write(sheetline, 14, "Red Easy", style3)

                if datared1 == teamloop or datared2 == teamloop or datared3 == teamloop:
                    if int(dataredsc) > int(databluesc):
                        sheet3.write(sheetline, 13, "Win", styleb)
                        win += 1
                    else:
                        sheet3.write(sheetline, 13, "Lose", styleblank)

                elif datablue1 == teamloop or datablue2 == teamloop or datablue3 == teamloop:
                    if int(dataredsc) < int(databluesc):
                        sheet3.write(sheetline, 13, "Win", styleb)
                        win += 1
                    else:
                        sheet3.write(sheetline, 13, "Lose", styleblank)

                # To see if 0 = 0

                if int(dataredsc) == 0 and int(databluesc) == 0:
                    sheetline -= 1
                    matches -= 1

                elif int(dataredsc) == 0:
                    sheet3.write(sheetline, 15, "Red DQ?", styleblank)
                elif int(databluesc) == 0:
                    sheet3.write(sheetline, 15, "Blue DQ?", styleblank)

                sheetline += 1
                loop += 1

                if loop > 2:
                    break

                output.append(line)

                # pprint.pprint(output)

                time.sleep(sleeptimer)

            sheetline += 1

            teaminfoline += 1

            decimal = (int(win) / int(matches))
            flag = decimal * 100
            flag = Decimal.from_float(flag).quantize(Decimal('0.0'))

            if float(flag) >= 70:
                sheet3.write(teaminfoline, 8, str(flag) + "%", style70)
                for x in range(9, 21):
                    sheet3.write(teaminfoline, x, "", style70)

            elif float(flag) >= 50:
                sheet3.write(teaminfoline, 8, str(flag) + "%", style50)
                for x in range(9, 21):
                    sheet3.write(teaminfoline, x, "", style50)

            elif float(flag) >= 30:
                sheet3.write(teaminfoline, 8, str(flag) + "%", style30)
                for x in range(9, 21):
                    sheet3.write(teaminfoline, x, "", style30)

            else:
                sheet3.write(teaminfoline, 8, str(flag) + "%", style0)
                for x in range(9, 21):
                    sheet3.write(teaminfoline, x, "", style0)
            for x in range(0, 21):
                sheet3.write(sheetline, x, "- - - - - - -", styleblank)

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


def excelgetallbugs():  # 204
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

            from urllib.request import urlopen
            r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                        teamloop + '&season=Turning%20Point')
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
                sheet10.write(sheetline, 6, "Positive", style1)
            elif int(datawins) < int(datalosses):
                sheet10.write(sheetline, 6, "Negative", style2)

            sheetline += 1

            # pprint.pprint(output)

            from urllib.request import urlopen

            r = urlopen('https://api.vexdb.io/v1/get_matches?team=' +
                        teamloop + '&season=Turning%20Point')

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
                sheet10.write(sheetline, 2, datared1, stylered)
                sheet10.write(sheetline, 3, datared2, stylered)
                sheet10.write(sheetline, 4, datared3, stylered)
                sheet10.write(sheetline, 5, dataredsit, stylered)
                sheet10.write(sheetline, 6, datablue1, styleblue)
                sheet10.write(sheetline, 7, datablue2, styleblue)
                sheet10.write(sheetline, 8, datablue3, styleblue)
                sheet10.write(sheetline, 9, databluesit, styleblue)
                sheet10.write(sheetline, 10, dataredsc, stylered)
                sheet10.write(sheetline, 11, databluesc, styleblue)
                sheet10.write(sheetline, 12, datateam + " =", styleb)

                if int(dataredsc) > int(databluesc):
                    sheet10.write(sheetline, 14, "Red", style1)
                elif int(dataredsc) < int(databluesc):
                    sheet10.write(sheetline, 14, "Blue", style2)

                if int(dataredsc) + 20 < int(databluesc):
                    sheet10.write(sheetline, 14, "Blue Easy", style4)
                elif int(dataredsc) - 20 > int(databluesc):
                    sheet10.write(sheetline, 14, "Red Easy", style3)

                if datared1 == teamloop or datared2 == teamloop or datared3 == teamloop:
                    if int(dataredsc) > int(databluesc):
                        sheet10.write(sheetline, 13, "Win", styleb)
                        win += 1
                    else:
                        sheet10.write(sheetline, 13, "Lose", styleblank)

                elif datablue1 == teamloop or datablue2 == teamloop or datablue3 == teamloop:
                    if int(dataredsc) < int(databluesc):
                        sheet10.write(sheetline, 13, "Win", styleb)
                        win += 1
                    else:
                        sheet10.write(sheetline, 13, "Lose", styleblank)

                # To see if 0 = 0

                if int(dataredsc) == 0 and int(databluesc) == 0:
                    sheetline -= 1
                    matches -= 1

                elif int(dataredsc) == 0:
                    sheet10.write(sheetline, 15, "Red DQ?", styleblank)
                elif int(databluesc) == 0:
                    sheet10.write(sheetline, 15, "Blue DQ?", styleblank)

                sheetline += 1
                loop += 1

                if loop > 2:
                    break

                output.append(line)

                # pprint.pprint(output)

                time.sleep(sleeptimer)

            sheetline += 1

            teaminfoline += 1

            decimal = (int(win) / int(matches))
            flag = decimal * 100
            flag = Decimal.from_float(flag).quantize(Decimal('0.0'))

            if float(flag) >= 70:
                sheet10.write(teaminfoline, 8, str(flag) + "%", style70)
                for x in range(9, 21):
                    sheet10.write(teaminfoline, x, "", style70)

            elif float(flag) >= 50:
                sheet10.write(teaminfoline, 8, str(flag) + "%", style50)
                for x in range(9, 21):
                    sheet10.write(teaminfoline, x, "", style50)

            elif float(flag) >= 30:
                sheet10.write(teaminfoline, 8, str(flag) + "%", style30)
                for x in range(9, 21):
                    sheet10.write(teaminfoline, x, "", style30)

            else:
                sheet10.write(teaminfoline, 8, str(flag) + "%", style0)
                for x in range(9, 21):
                    sheet10.write(teaminfoline, x, "", style0)
            for x in range(0, 21):
                sheet10.write(sheetline, x, "- - - - - - -", styleblank)

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


def excelgetweneed():  # 205
    time.sleep(1)
    number = 0
    sheetline = 0
    start = time.time()
    list1 = ['2U', '81K', '169A', '365X', '624K', '934Z', '1064A', '1437Z', '1961U', '2360S', '2719B', '3269B', '3767A',
             '4815B', '5139A', '6627A', '6741A', '7258B', '7536B', '7853A', '8110B', '8192B', '9060C', '9228A', '9551B',
             '9932E', '10955M', '17071B', '35211C', '97934U', '98807A']

    while True:

        while number < int(len(list1)):

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

            from urllib.request import urlopen
            r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                        teamloop + '&season=Turning%20Point')
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
                sheet6.write(sheetline, 6, "Positive", style1)
            elif int(datawins) < int(datalosses):
                sheet6.write(sheetline, 6, "Negative", style2)

            sheetline += 1

            # pprint.pprint(output)

            from urllib.request import urlopen

            r = urlopen('https://api.vexdb.io/v1/get_matches?team=' +
                        teamloop + '&season=Turning%20Point')

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
                sheet6.write(sheetline, 2, datared1, stylered)
                sheet6.write(sheetline, 3, datared2, stylered)
                sheet6.write(sheetline, 4, datared3, stylered)
                sheet6.write(sheetline, 5, dataredsit, stylered)
                sheet6.write(sheetline, 6, datablue1, styleblue)
                sheet6.write(sheetline, 7, datablue2, styleblue)
                sheet6.write(sheetline, 8, datablue3, styleblue)
                sheet6.write(sheetline, 9, databluesit, styleblue)
                sheet6.write(sheetline, 10, dataredsc, stylered)
                sheet6.write(sheetline, 11, databluesc, styleblue)
                sheet6.write(sheetline, 12, datateam + " =", styleb)

                if int(dataredsc) > int(databluesc):
                    sheet6.write(sheetline, 14, "Red", style1)
                elif int(dataredsc) < int(databluesc):
                    sheet6.write(sheetline, 14, "Blue", style2)

                if int(dataredsc) + 20 < int(databluesc):
                    sheet6.write(sheetline, 14, "Blue Easy", style4)
                elif int(dataredsc) - 20 > int(databluesc):
                    sheet6.write(sheetline, 14, "Red Easy", style3)

                if datared1 == teamloop or datared2 == teamloop or datared3 == teamloop:
                    if int(dataredsc) > int(databluesc):
                        sheet6.write(sheetline, 13, "Win", styleb)
                        win += 1
                    else:
                        sheet6.write(sheetline, 13, "Lose", styleblank)

                elif datablue1 == teamloop or datablue2 == teamloop or datablue3 == teamloop:
                    if int(dataredsc) < int(databluesc):
                        sheet6.write(sheetline, 13, "Win", styleb)
                        win += 1
                    else:
                        sheet6.write(sheetline, 13, "Lose", styleblank)

                # To see if 0 = 0

                if int(dataredsc) == 0 and int(databluesc) == 0:
                    sheetline -= 1
                    matches -= 1

                elif int(dataredsc) == 0:
                    sheet6.write(sheetline, 15, "Red DQ?", styleblank)
                elif int(databluesc) == 0:
                    sheet6.write(sheetline, 15, "Blue DQ?", styleblank)

                sheetline += 1
                loop += 1

                if loop > 2:
                    break

                output.append(line)

                # pprint.pprint(output)

                time.sleep(sleeptimer)

            sheetline += 1

            teaminfoline += 1

            decimal = (int(win) / int(matches))
            flag = decimal * 100
            flag = Decimal.from_float(flag).quantize(Decimal('0.0'))

            if float(flag) >= 70:
                sheet6.write(teaminfoline, 8, str(flag) + "%", style70)
                for x in range(9, 21):
                    sheet6.write(teaminfoline, x, "", style70)

            elif float(flag) >= 50:
                sheet6.write(teaminfoline, 8, str(flag) + "%", style50)
                for x in range(9, 21):
                    sheet6.write(teaminfoline, x, "", style50)

            elif float(flag) >= 30:
                sheet6.write(teaminfoline, 8, str(flag) + "%", style30)
                for x in range(9, 21):
                    sheet6.write(teaminfoline, x, "", style30)

            else:
                sheet6.write(teaminfoline, 8, str(flag) + "%", style0)
                for x in range(9, 21):
                    sheet6.write(teaminfoline, x, "", style0)

            for x in range(0, 21):
                sheet6.write(sheetline, x, "- - - - - - -", styleblank)

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


def excelscanworld():
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

            from urllib.request import urlopen
            r = urlopen(
                'https://api.vexdb.io/v1/get_rankings?team=' + teamloop + '&season=Turning%20Point' + '&sku=RE-VRC-17'
                                                                                                      '-3805')
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
                sheet5.write(sheetline, 6, "Positive", style1)
            elif int(datawins) < int(datalosses):
                sheet5.write(sheetline, 6, "Negative", style2)

            sheetline += 1

            # pprint.pprint(output)

            from urllib.request import urlopen

            r = urlopen('https://api.vexdb.io/v1/get_matches?team=' +
                        teamloop + '&season=Turning%20Point')

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
                sheet5.write(sheetline, 2, datared1, stylered)
                sheet5.write(sheetline, 3, datared2, stylered)
                sheet5.write(sheetline, 4, datared3, stylered)
                sheet5.write(sheetline, 5, dataredsit, stylered)
                sheet5.write(sheetline, 6, datablue1, styleblue)
                sheet5.write(sheetline, 7, datablue2, styleblue)
                sheet5.write(sheetline, 8, datablue3, styleblue)
                sheet5.write(sheetline, 9, databluesit, styleblue)
                sheet5.write(sheetline, 10, dataredsc, stylered)
                sheet5.write(sheetline, 11, databluesc, styleblue)
                sheet5.write(sheetline, 12, datateam + " =", styleb)

                if int(dataredsc) > int(databluesc):
                    sheet5.write(sheetline, 14, "Red", style1)
                elif int(dataredsc) < int(databluesc):
                    sheet5.write(sheetline, 14, "Blue", style2)

                if int(dataredsc) + 20 < int(databluesc):
                    sheet5.write(sheetline, 14, "Blue Easy", style4)
                elif int(dataredsc) - 20 > int(databluesc):
                    sheet5.write(sheetline, 14, "Red Easy", style3)

                if datared1 == teamloop or datared2 == teamloop or datared3 == teamloop:
                    if int(dataredsc) > int(databluesc):
                        sheet5.write(sheetline, 13, "Win", styleb)
                    else:
                        sheet5.write(sheetline, 13, "Lose", styleblank)
                elif datablue1 == teamloop or datablue2 == teamloop or datablue3 == teamloop:
                    if int(dataredsc) < int(databluesc):
                        sheet5.write(sheetline, 13, "Win", styleb)
                    else:
                        sheet5.write(sheetline, 13, "Lose", styleblank)

                sheetline += 1
                loop += 1

                if loop > 2:
                    break

                output.append(line)

                # pprint.pprint(output)

                time.sleep(0.1)

            sheetline += 1
            for x in range(0, 15):
                sheet5.write(sheetline, x, "- - - - - - -", styleblank)

            sheetline += 1
            for x in range(0, 15):
                sheet5.write(sheetline, x, "- - - - - - -", styleblank)
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


def excelteammatches():
    name = input('Team #?\n')
    print('Checking, TEAM %s.' % name)

    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_matches?team=' +
                name + '&season=Turning%20Point')
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


def searchteamcurrentseason():
    name = str(input('Team #?\n'))
    print('Checking, TEAM %s.' % name)
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                name + '&season=Turning%20Point')
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


def getalldata():
    # getalldata
    print("This will show the recent three matches.")
    name = str(input('Team #?\n'))
    print('Checking, TEAM %s.' % name)
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                name + '&season=Turning%20Point')
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
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_matches?team=' +
                name + '&season=Turning%20Point')
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

def empty():  # will remove
    print("Nothing here yet.")
    print("Return to Menu.")
    print()
    time.sleep(1)


def timeisout():
    # Input Team
    global_var.inputmode = str(
        input("Type in the preset value or 6 teams separate by ,\n"))
    teams = ""

        # global_var.teamr1,global_var.teamr2,global_var.teamr3,global_var.teamb1,global_var.teamb2,global_var.teamb3 = input("Please input 6 teams separate by ,\n").split(',')
    print(
        "TR1: " + global_var.teamr1 + " TR2: " + global_var.teamr2 + " TR3: " + global_var.teamr3 + " || TB1: "
        + global_var.teamb1 + " TB2: " + global_var.teamb2 + " TB3: " + global_var.teamb3)

    if str(global_var.teamr1) != "":
        global_var.teamsent = global_var.teamr1
        global_var.teamname = global_var.teamr1
        teamskill()
        global_var.teamr1skillout = global_var.skillave
        global_var.teamr1wins = global_var.winsave
        global_var.teamr1ap = global_var.apave
        global_var.teamr1ranking = global_var.rankave
        global_var.teamr1highest = global_var.highestave
        global_var.teamr1ccwm = global_var.ccwmave
        global_var.teamr1dpr = global_var.dprave
        global_var.teamr1opr = global_var.oprave
        global_var.teamr1currentranking = global_var.currentranking
        global_var.teamr1currentwins = global_var.currentwins
        global_var.teamr1currentlosses = global_var.currentlosses
    else:
        print("Team Red 1 is blank.")

    if str(global_var.teamr2) != "":
        global_var.teamsent = global_var.teamr2
        global_var.teamname = global_var.teamr2
        teamskill()
        global_var.teamr2skillout = global_var.skillave
        global_var.teamr2wins = global_var.winsave
        global_var.teamr2ap = global_var.apave
        global_var.teamr2ranking = global_var.rankave
        global_var.teamr2highest = global_var.highestave
        global_var.teamr2ccwm = global_var.ccwmave
        global_var.teamr2dpr = global_var.dprave
        global_var.teamr2opr = global_var.oprave
        global_var.teamr2currentranking = global_var.currentranking
        global_var.teamr2currentwins = global_var.currentwins
        global_var.teamr2currentlosses = global_var.currentlosses
    else:
        print("Team Red 2 is blank.")

    if str(global_var.teamr3) != "":
        global_var.teamsent = global_var.teamr3
        global_var.teamname = global_var.teamr3
        teamskill()
        global_var.teamr3skillout = global_var.skillave
        global_var.teamr3wins = global_var.winsave
        global_var.teamr3ap = global_var.apave
        global_var.teamr3ranking = global_var.rankave
        global_var.teamr3highest = global_var.highestave
        global_var.teamr3ccwm = global_var.ccwmave
        global_var.teamr3dpr = global_var.dprave
        global_var.teamr3opr = global_var.oprave
        global_var.teamr3currentranking = global_var.currentranking
        global_var.teamr3currentwins = global_var.currentwins
        global_var.teamr3currentlosses = global_var.currentlosses
    else:
        print("Team Red 3 is blank.")

    if str(global_var.teamb1) != "":
        global_var.teamsent = global_var.teamb1
        global_var.teamname = global_var.teamb1
        teamskill()
        global_var.teamb1skillout = global_var.skillave
        global_var.teamb1wins = global_var.winsave
        global_var.teamb1ap = global_var.apave
        global_var.teamb1ranking = global_var.rankave
        global_var.teamb1highest = global_var.highestave
        global_var.teamb1ccwm = global_var.ccwmave
        global_var.teamb1dpr = global_var.dprave
        global_var.teamb1opr = global_var.oprave
        global_var.teamb1currentranking = global_var.currentranking
        global_var.teamb1currentwins = global_var.currentwins
        global_var.teamb1currentlosses = global_var.currentlosses
    else:
        print("Team Blue 1 is blank.")

    if str(global_var.teamb2) != "":
        global_var.teamsent = global_var.teamb2
        global_var.teamname = global_var.teamb2
        teamskill()
        global_var.teamb2skillout = global_var.skillave
        global_var.teamb2wins = global_var.winsave
        global_var.teamb2ap = global_var.apave
        global_var.teamb2ranking = global_var.rankave
        global_var.teamb2highest = global_var.highestave
        global_var.teamb2ccwm = global_var.ccwmave
        global_var.teamb2dpr = global_var.dprave
        global_var.teamb2opr = global_var.oprave
        global_var.teamb2currentranking = global_var.currentranking
        global_var.teamb2currentwins = global_var.currentwins
        global_var.teamb2currentlosses = global_var.currentlosses
    else:
        print("Team Blue 2 is blank.")

    if str(global_var.teamb3) != "":
        global_var.teamsent = global_var.teamb3
        global_var.teamname = global_var.teamb3
        teamskill()
        global_var.teamb3skillout = global_var.skillave
        global_var.teamr3wins = global_var.winsave
        global_var.teamb3ap = global_var.apave
        global_var.teamb3ranking = global_var.rankave
        global_var.teamb3highest = global_var.highestave
        global_var.teamb3ccwm = global_var.ccwmave
        global_var.teamb3dpr = global_var.dprave
        global_var.teamb3opr = global_var.oprave
        global_var.teamb3currentranking = global_var.currentranking
        global_var.teamb3currentwins = global_var.currentwins
        global_var.teamb3currentlosses = global_var.currentlosses
    else:
        print("Team Blue 3 is blank.")

    # print("Skill is average of all this season. Auto is the previous competition. (Should be the state final)")
    # print("Ranking is (10-Ranking), if the team is not the first 10th, it will show as 0.")

    graphbubble()  # pass value use arg instead of global

    return None


# global_var.teamr1skillout,global_var.teamr2skillout,global_var.teamr3skillout,global_var.teamb1skillout,global_var.teamb2skillout,global_var.teamb3skillout
'''
def teamskill():
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_skills?team=' + global_var.teamsent + '&season=Turning%20Point&type=2')
    text = r.read()

    json_dict = json.loads(text)

    output = []
    skilltotal = 0
    totalattempts = 0

    for r in json_dict["result"]:
        skill = int(r["score"])
        attempt = int(r["attempts"])
        if int(attempt) != 0 and str(attempt) != "":
            totalattempts += 1
        skilltotal = skill + skilltotal
        global_var.skillave = float(skilltotal) / int(attempt)
    ''' '''
    if int(totalattempts) != 0:
        skillave = int(skilltotal) / int(totalattempts)
        print(global_var.teamname + ": " + str(global_var.skillave))
        teamsent()
    ''' '''
    teamsent()
'''


def teamskill():
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_skills?team=' +
                global_var.teamsent + '&season=Turning%20Point')
    text = r.read()
    json_dict = json.loads(text)
    # output = []
    skilltotal = 0
    totalattempts = 0
    skillave = 0

    for r in json_dict["result"]:
        skill = int(r["score"])
        attempt = int(r["attempts"])
        if int(attempt) != 0:
            totalattempts += 1
        skilltotal = skill + skilltotal

    if int(totalattempts) != 0:
        skillave = int(skilltotal) / int(totalattempts)
    else:
        skillave = 0

    decimal = skillave
    decimal = Decimal.from_float(decimal).quantize(Decimal('0.0'))
    global_var.skillave = decimal
    teamskill = float(global_var.skillave)
    print(global_var.teamname + ": " + str(global_var.skillave))
    teamsent()


def teamsent():
    count = 0
    global_var.winsave = 0
    teamwins = 0
    global_var.winstotal = 0
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                global_var.teamsent + '&season=Turning%20Point')
    text = r.read()
    json_dict = json.loads(text)
    for r in json_dict["result"]:
        # line = '{}'.format(r["wins"])
        teamwins = '{}'.format(r["wins"])
        count += 1
        global_var.winstotal = teamwins + teamwins
        if teamwins == "" or teamwins == "":
            print("break cuz blank")
            count -= 1
            global_var.winsave = float(global_var.winstotal) / int(count)
            teamap()
        global_var.winsave = float(global_var.winstotal) / int(count)
    teamcurrent()


def teamcurrent():  # can be part of teamsent()
    global_var.currentranking = 0
    global_var.currentwins = 0
    global_var.currentlosses = 0
    from urllib.request import urlopen
    r = urlopen(
        'https://api.vexdb.io/v1/get_rankings?team=' + global_var.teamsent + '&season=Turning%20Point'
        + global_var.CONST_match)
    text = r.read()
    json_dict = json.loads(text)
    for r in json_dict["result"]:
        # line = '{}'.format(r["rank"], r["wins"], r["losses"])
        # output.append(line)
        global_var.currentranking = '{}'.format(r["rank"])
        global_var.currentwins = '{}'.format(r["wins"])
        global_var.currentlosses = '{}'.format(r["losses"])
    teamap()


def teamap():
    teammap = 0
    global_var.aptotal = 0
    count = 0
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                global_var.teamsent + '&season=Turning%20Point')
    text = r.read()
    json_dict = json.loads(text)

    for r in json_dict["result"]:
        # line = '{}'.format(r["ap"])
        # output.append(line)
        teammap = '{}'.format(r["ap"])
        count += 1
        diff = 0

        if int(teammap) > 25:
            diff = (int(teammap) - 25) * 0.2
            teammap = 25 + float(diff)
            print("Balance over 25, " + str(diff))
        global_var.aptotal = int(global_var.aptotal) + int(teammap)
        global_var.apave = int(global_var.aptotal) / int(count)

        if teammap == "" or teammap == "":
            print("break cuz blank")
            count -= 1
            teammap = global_var.apave
            teamranking()
    teamranking()


def teamranking():
    TeamRanking = 0
    global_var.ranktotal = 0
    global_var.rankave = 0
    count = 0
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                global_var.teamsent + '&season=Turning%20Point')
    text = r.read()
    json_dict = json.loads(text)
    for r in json_dict["result"]:
        # line = '{}'.format(r["rank"])
        # output.append(line)
        TeamRanking = '{}'.format(r["rank"])
        count += 1
        global_var.ranktotal = int(
            global_var.ranktotal) + int(TeamRanking)
        global_var.rankave = float(global_var.ranktotal) / int(count)

        if TeamRanking == "" or TeamRanking == "":
            print("break cuz blank")
            count -= 1
            global_var.rankave = float(global_var.ranktotal) / int(count)
            teamhighest()
        global_var.rankave = float(TeamRanking) / int(count)
    teamhighest()


def teamhighest():
    highesttotal = 0
    global_var.highestave = 0
    count = 0
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                global_var.teamsent + '&season=Turning%20Point')
    text = r.read()
    json_dict = json.loads(text)
    for r in json_dict["result"]:
        # line = '{}'.format(r["max_score"])
        # output.append(line)
        TeamHighest = '{}'.format(r["max_score"])
        count += 1
        highesttotal = int(
            highesttotal) + int(TeamHighest)
        global_var.highestave = int(highesttotal) / count
        if TeamHighest == "":
            print("break cuz blank")
            count -= 1
            global_var.highestave = float(highesttotal) / int(count)
            teampr()
        global_var.highestave = float(highesttotal) / int(count)
    teampr()


def teampr():
    global_var.oprtotal = 0
    global_var.dprtotal = 0
    teamopr = 0
    teamdpr = 0
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                global_var.teamsent + '&season=Turning%20Point')
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
        global_var.oprtotal = float(
            global_var.oprtotal) + float(teamopr)
        global_var.oprave = float(global_var.oprtotal) / int(count)
        global_var.dprtotal = float(
            global_var.dprtotal) + float(teamdpr)
        global_var.dprave = float(global_var.dprtotal) / int(count)

        if teamdpr == "" or teamopr == "":
            print("break cuz blank")
            count -= 1
            teamdpr = float(global_var.dprave)
            teamopr = float(global_var.oprave)
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
    teamccwm = 0
    ccwmtotal = 0
    global_var.ccwmave = 0
    from urllib.request import urlopen
    r = urlopen('https://api.vexdb.io/v1/get_rankings?team=' +
                global_var.teamsent + '&season=Turning%20Point')
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
        global_var.ccwmave = float(ccwmtotal) / int(count)
        if teamccwm == "" or teamccwm == "":
            print("break cuz blank")
            count -= 18
            teamccwm = float(global_var.ccwmave)
            break
        teamccwm = float(global_var.ccwmave)


def graphbubble():  # it should be part of "timeisout"
    global_var.teamr1skillout = float(global_var.teamr1skillout) / 10
    global_var.teamr2skillout = float(global_var.teamr2skillout) / 10
    global_var.teamr3skillout = float(global_var.teamr3skillout) / 10
    global_var.teamr1ap = round(float(global_var.teamr1ap) / 5, 1)
    global_var.teamr2ap = round(float(global_var.teamr2ap) / 5, 1)
    global_var.teamr3ap = round(float(global_var.teamr3ap) / 5, 1)
    # The Formula
    global_var.teamr1ranking = int(10 - int(global_var.teamr1ranking))
    global_var.teamr2ranking = int(10 - int(global_var.teamr2ranking))
    global_var.teamr3ranking = int(10 - int(global_var.teamr3ranking))

    # /17
    global_var.teamr1highest = round(
        float(int(global_var.teamr1highest) / 17), 1)
    global_var.teamr2highest = round(
        float(int(global_var.teamr2highest) / 17), 1)
    global_var.teamr3highest = round(
        float(int(global_var.teamr3highest) / 17), 1)

    if int(global_var.teamr1ranking) < 0:
        global_var.teamr1ranking = 0
    if int(global_var.teamr2ranking) < 0:
        global_var.teamr2ranking = 0
    if int(global_var.teamr3ranking) < 0:
        global_var.teamr3ranking = 0

    # Check
    print("Skill " + str(global_var.teamr1skillout) + " " + str(global_var.teamr2skillout) + " " + str(
        global_var.teamr3skillout))
    print("Season Wins " + str(global_var.teamr1wins) + " " + str(global_var.teamr2wins) + " " + str(
        global_var.teamr3wins))
    print("AP " + str(global_var.teamr1ap) + " " +
          str(global_var.teamr2ap) + " " + str(global_var.teamr3ap))
    print("Ranking " + str(global_var.teamr1ranking) + " " + str(global_var.teamr2ranking) + " " + str(
        global_var.teamr3ranking))
    print("Highest " + str(global_var.teamr1highest) + " " + str(global_var.teamr2highest) + " " + str(
        global_var.teamr3highest))
    print("CCWM" + str(global_var.teamr1ccwm))

    global_var.teamb1skillout = float(global_var.teamb1skillout) / 10
    global_var.teamb2skillout = float(global_var.teamb2skillout) / 10
    global_var.teamb3skillout = float(global_var.teamb3skillout) / 10

    global_var.teamb1ap = round(float(global_var.teamb1ap) / 5, 1)
    global_var.teamb2ap = round(float(global_var.teamb2ap) / 5, 1)
    global_var.teamb3ap = round(float(global_var.teamb3ap) / 5, 1)

    # The Formula
    global_var.teamb1ranking = int(10 - int(global_var.teamb1ranking))
    global_var.teamb2ranking = int(10 - int(global_var.teamb2ranking))
    global_var.teamb3ranking = int(10 - int(global_var.teamb3ranking))

    # /17
    global_var.teamb1highest = round(
        float(int(global_var.teamb1highest) / 17), 1)
    global_var.teamb2highest = round(
        float(int(global_var.teamb2highest) / 17), 1)
    global_var.teamb3highest = round(
        float(int(global_var.teamb3highest) / 17), 1)

    if int(global_var.teamb1ranking) <= 0:
        global_var.teamb1ranking = 0
    if int(global_var.teamb2ranking) <= 0:
        global_var.teamb2ranking = 0
    if int(global_var.teamb3ranking) <= 0:
        global_var.teamb3ranking = 0

    # Check
    print("Skill " + str(global_var.teamb1skillout) + " " + str(global_var.teamb2skillout) + " " + str(
        global_var.teamb3skillout))
    print("Season Wins " + str(global_var.teamb1wins) + " " + str(global_var.teamb2wins) + " " + str(
        global_var.teamb3wins))
    print("AP " + str(global_var.teamb1ap) + " " +
          str(global_var.teamb2ap) + " " + str(global_var.teamb3ap))
    print("Ranking " + str(global_var.teamb1ranking) + " " + str(global_var.teamb2ranking) + " " + str(
        global_var.teamb3ranking))
    print("Highest " + str(global_var.teamb1highest) + " " + str(global_var.teamb2highest) + " " + str(
        global_var.teamb3highest))

    if global_var.teamr1ccwm < 0:
        global_var.teamr1ccwm = 0.1
    if global_var.teamr2ccwm < 0:
        global_var.teamr2ccwm = 0.1
    if global_var.teamr3ccwm < 0:
        global_var.teamr3ccwm = 0.1
    if global_var.teamb1ccwm < 0:
        global_var.teamb1ccwm = 0.1
    if global_var.teamb2ccwm < 0:
        global_var.teamb2ccwm = 0.1
    if global_var.teamb3ccwm < 0:
        global_var.teamb3ccwm = 0.1

    # create data!

    x = float(global_var.teamr1skillout)
    y = float(global_var.teamr1ap)
    # z = float(global_var.teamr1wins)
    z = float(global_var.teamr1highest)
    plt.text(x, y, str(global_var.teamr1), ha='center',
             va='center', fontweight='bold', color='red')
    plt.scatter(x, y, s=z * 300, c="red", alpha=0.4, linewidth=6)

    x = float(global_var.teamr2skillout)
    y = float(global_var.teamr2ap)
    # z = float(global_var.teamr2wins)
    z = float(global_var.teamr2highest)
    plt.text(x, y, str(global_var.teamr2), ha='center',
             va='center', fontweight='bold', color='red')
    plt.scatter(x, y, s=z * 300, c="red", alpha=0.4, linewidth=6)

    x = float(global_var.teamr3skillout)
    y = float(global_var.teamr3ap)
    # z = float(global_var.teamr3wins)
    z = float(global_var.teamr3highest)
    plt.text(x, y, str(global_var.teamr3), ha='center',
             va='center', fontweight='bold', color='red')
    plt.scatter(x, y, s=z * 300, c="red", alpha=0.4, linewidth=6)

    x = float(global_var.teamr1dpr)
    y = float(global_var.teamr1opr)
    # z = float(global_var.teamr1wins)
    z = float(global_var.teamr1ccwm)
    plt.text(x, y, str("[" + global_var.teamr1 + "]"), ha='center',
             fontweight='bold', va='center', color='darkred')
    plt.scatter(x, y, s=z * 50, c="deeppink", alpha=0.4, linewidth=6)

    x = float(global_var.teamr2dpr)
    y = float(global_var.teamr2opr)
    # z = float(global_var.teamr2wins)
    z = float(global_var.teamr2ccwm)
    plt.text(x, y, str("[" + global_var.teamr2 + "]"), ha='center',
             fontweight='bold', va='center', color='darkred')
    plt.scatter(x, y, s=z * 50, c="deeppink", alpha=0.4, linewidth=6)

    if global_var.teamr3dpr != 0:
        x = float(global_var.teamr3dpr)
        y = float(global_var.teamr3opr)
        # z = float(global_var.teamr3wins)
        z = float(global_var.teamr3ccwm)
        plt.text(x, y, str("[" + global_var.teamr3 + "]"), ha='center',
                 fontweight='bold', va='center', color='darkred')
        plt.scatter(x, y, s=z * 50, c="deeppink", alpha=0.4, linewidth=6)

    x = float(global_var.teamb1skillout)
    y = float(global_var.teamb1ap)
    # z = float(global_var.teamb1wins)
    z = float(global_var.teamb1highest)
    plt.text(x, y, str(global_var.teamb1), ha='center',
             va='center', fontweight='bold', color='royalblue')
    plt.scatter(x, y, s=z * 300, c="royalblue", alpha=0.4, linewidth=6)

    x = float(global_var.teamb2skillout)
    y = float(global_var.teamb2ap)
    # z = float(global_var.teamb2wins)
    z = float(global_var.teamb2highest)
    plt.text(x, y, str(global_var.teamb2), ha='center',
             va='center', fontweight='bold', color='royalblue')
    plt.scatter(x, y, s=z * 300, c="royalblue", alpha=0.4, linewidth=6)

    x = float(global_var.teamb3skillout)
    y = float(global_var.teamb3ap)
    # z = float(global_var.teamb3wins)
    z = float(global_var.teamb3highest)
    plt.text(x, y, str(global_var.teamb3), ha='center',
             va='center', fontweight='bold', color='royalblue')
    plt.scatter(x, y, s=z * 300, c="royalblue", alpha=0.4, linewidth=6)

    x = float(global_var.teamb1dpr)
    y = float(global_var.teamb1opr)
    # z = float(global_var.teamb1wins)
    z = float(global_var.teamb1ccwm)
    plt.text(x, y, str("[" + global_var.teamb1 + "]"), ha='center',
             va='bottom', fontweight='bold', color='dodgerblue')
    plt.scatter(x, y, s=z * 50, c="dodgerblue", alpha=0.4, linewidth=6)

    x = float(global_var.teamb2dpr)
    y = float(global_var.teamb2opr)
    # z = float(global_var.teamb2wins)
    z = float(global_var.teamb2ccwm)
    plt.text(x, y, str("[" + global_var.teamb2 + "]"), ha='center',
             va='bottom', fontweight='bold', color='dodgerblue')
    plt.scatter(x, y, s=z * 50, c="dodgerblue", alpha=0.4, linewidth=6)

    if global_var.teamb3dpr != 0:
        x = float(global_var.teamb3dpr)
        y = float(global_var.teamb3opr)
        # z = float(global_var.teamb3wins)
        z = float(global_var.teamb3ccwm)
        plt.text(x, y, str("[" + global_var.teamb3 + "]"), ha='center', va='bottom', fontweight='bold',
                 color='dodgerblue')
        plt.scatter(x, y, s=z * 50, c="dodgerblue", alpha=0.4, linewidth=6)

    xmin, xmax = plt.xlim()
    ymin, ymax = plt.ylim()
    xaxis = float(xmax)
    xmiddle = (float(xaxis) / 2)
    # Add titles (main and on axis)
    try:
        os.remove("graph/" + global_var.inputmode + ".png")
        print("Previous deleted.")
        time.sleep(1)
    except OSError:
        print("something is not right")
        pass
    plt.xlabel(
        "Skill / [Defensive]")
    plt.ylabel("AP / [Offensive]")
    plt.title(
        "Red: " + global_var.teamr1 + " " + global_var.teamr2 + " " + global_var.teamr3 +
        " Blue: " + global_var.teamb1 + " " +
        global_var.teamb2 + " " + global_var.teamb3,
        loc="left")
    plt.text(xmiddle, -0.02,
             "Team #, X: Skill, Y: AP, Z: Highest Score\n [Team #], X: Defensive Pts Y: Offensive Pts Z: Contribution",
             ha='center', color='white', bbox=dict(facecolor='darkslateblue', alpha=0.5))
    plt.text((xmin + 0.3), (ymax - 0.5), global_var.teamr1 + " W: " + str(global_var.teamr1currentwins) + " L: " + str(
        global_var.teamr1currentlosses) + " R: " + str(
        global_var.teamr1currentranking) + "\n" + global_var.teamr2 + " W: " + str(
        global_var.teamr2currentwins) + " L: " + str(global_var.teamr2currentlosses) + " R: " + str(
        global_var.teamr2currentranking) + "\n" + global_var.teamr3 + " W: " + str(
        global_var.teamr3currentwins) + " L: " + str(global_var.teamr3currentlosses) + " R: " + str(
        global_var.teamr3currentranking) + "\n" + global_var.teamb1 + " W: " + str(
        global_var.teamb1currentwins) + " L: " + str(global_var.teamb1currentlosses) + " R: " + str(
        global_var.teamb1currentranking) + "\n" + global_var.teamb2 + " W: " + str(
        global_var.teamb2currentwins) + " L: " + str(global_var.teamb2currentlosses) + " R: " + str(
        global_var.teamb2currentranking) + "\n" + global_var.teamb3 + " W: " + str(
        global_var.teamb3currentwins) + " L: " + str(global_var.teamb3currentlosses) + " R: " + str(
        global_var.teamb3currentranking), ha='left', va='top', color='white', fontsize='smaller',
             bbox=dict(facecolor='darkgreen', alpha=0.5))
    plt.savefig("graph/" + global_var.inputmode + ".png")
    print("Graph poped and saved.")
    plt.show()


def answer():
    # answerr = 0
    # answerb = 0

    teamrexist = 0
    teambexist = 0

    # teamrskill = float(global_var.teamr1skillout) + float(global_var.teamr2skillout) + float(
    #     global_var.teamr3skillout)
    # teambskill = float(global_var.teamb1skillout) + float(global_var.teamb2skillout) + float(
    #     global_var.teamb3skillout)
    # teamrave = (float(global_var.teamrskill) / 3)
    # teambave = (float(global_var.teambskill)) / 3

    if global_var.teamr1skillout != 0:
        teamrexist += 1
    if global_var.teamr2skillout != 0:
        teamrexist += 1
    if global_var.teamr3skillout != 0:
        teamrexist += 1
    if global_var.teamb1skillout != 0:
        teambexist += 1
    if global_var.teamb2skillout != 0:
        teambexist += 1
    if global_var.teamb3skillout != 0:
        teambexist += 1

    time.sleep(2)
    input("Press Any Key to Continue\n")


def graphred():  # nothing use this
    # Set data
    fig1 = plt.figure('Red')

    global_var.teamr1skillout = float(global_var.teamr1skillout) / 10
    global_var.teamr2skillout = float(global_var.teamr2skillout) / 10
    global_var.teamr3skillout = float(global_var.teamr3skillout) / 10

    global_var.teamr1ap = round(float(global_var.teamr1ap) / 5, 1)
    global_var.teamr2ap = round(float(global_var.teamr2ap) / 5, 1)
    global_var.teamr3ap = round(float(global_var.teamr3ap) / 5, 1)

    # The Formula
    global_var.teamr1ranking = int(10 - int(global_var.teamr1ranking))
    global_var.teamr2ranking = int(10 - int(global_var.teamr2ranking))
    global_var.teamr3ranking = int(10 - int(global_var.teamr3ranking))

    # /17
    global_var.teamr1highest = round(float(int(global_var.teamr1highest) / 17), 1)
    global_var.teamr2highest = round(float(int(global_var.teamr2highest) / 17), 1)
    global_var.teamr3highest = round(float(int(global_var.teamr3highest) / 17), 1)

    if int(global_var.teamr1ranking) < 0:
        global_var.teamr1ranking = 0
    if int(global_var.teamr2ranking) < 0:
        global_var.teamr2ranking = 0
    if int(global_var.teamr3ranking) < 0:
        global_var.teamr3ranking = 0

    # Check
    print("Skill " + str(global_var.teamr1skillout) + " " + str(global_var.teamr2skillout) + " " + str(
        global_var.teamr3skillout))
    print("Season Wins " + str(global_var.teamr1wins) + " " + str(global_var.teamr2wins) + " " + str(
        global_var.teamr3wins))
    print("AP " + str(global_var.teamr1ap) + " " + str(global_var.teamr2ap) + " " + str(global_var.teamr3ap))
    print("Ranking " + str(global_var.teamr1ranking) + " " + str(global_var.teamr2ranking) + " " + str(
        global_var.teamr3ranking))
    print("Highest " + str(global_var.teamr1highest) + " " + str(global_var.teamr2highest) + " " + str(
        global_var.teamr3highest))

    df = pd.DataFrame.from_items([
        ('group', ['A', 'B', 'C', 'D']),
        ('Skill', [global_var.teamr1skillout, global_var.teamr2skillout, global_var.teamr3skillout, 0]),
        ('Season Wins', [global_var.teamr1wins, global_var.teamr2wins, global_var.teamr3wins, 0]),
        ('AP', [global_var.teamr1ap, global_var.teamr2ap, global_var.teamr3ap, 0]),  # /5
        ('Highest', [global_var.teamr1highest, global_var.teamr2highest, global_var.teamr3highest, 0]),  # /15
        ('Rankings', [global_var.teamr1ranking, global_var.teamr2ranking, global_var.teamr3ranking, 0])  # The Formula
    ])

    # 'Skill': [global_var.teamr1skillout, global_var.teamr2skillout, global_var.teamr3skillout,0],
    # 'Season Wins': [global_var.teamr1wins, global_var.teamr2wins, global_var.teamr3wins, 0],

    # ------- PART 1: Create background

    # number of variable
    categories = list(df)[1:]
    N = len(categories)

    # What will be the angle of each axis in the plot? (we divide the plot / number of variable)
    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]

    # Initialise the spider plot
    ax = plt.subplot(111, polar=True)

    # If you want the first axis to be on top:
    ax.set_theta_offset(pi / 2)
    ax.set_theta_direction(-1)

    # Draw one axe per variable + add labels labels yet
    plt.xticks(angles[:-1], categories)

    # Draw ylabels
    ax.set_rlabel_position(0)
    plt.yticks([3, 6, 9], ["3", "6", "9"], color="grey", size=7)
    plt.ylim(0, 12)

    # Ind1
    values = df.loc[0].drop('group').values.flatten().tolist()
    values += values[:1]
    ax.plot(angles, values, linewidth=1, linestyle='solid', label=str(global_var.teamr1))
    ax.fill(angles, values, 'b', alpha=0.1)

    # Ind2
    values = df.loc[1].drop('group').values.flatten().tolist()
    values += values[:1]
    ax.plot(angles, values, linewidth=1, linestyle='solid', label=str(global_var.teamr2))
    ax.fill(angles, values, 'r', alpha=0.1)

    # Ind3
    values = df.loc[2].drop('group').values.flatten().tolist()
    values += values[:1]
    ax.plot(angles, values, linewidth=1, linestyle='solid', label=str(global_var.teamr3))
    ax.fill(angles, values, 'r', alpha=0.1)

    # Add legend
    plt.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1))

    plt.show()
    plt.close()


def graphblue():
    # Set data
    fig2 = plt.figure('Blue')

    global_var.teamb1skillout = float(global_var.teamb1skillout) / 10
    global_var.teamb2skillout = float(global_var.teamb2skillout) / 10
    global_var.teamb3skillout = float(global_var.teamb3skillout) / 10

    global_var.teamb1ap = round(float(global_var.teamb1ap) / 5, 1)
    global_var.teamb2ap = round(float(global_var.teamb2ap) / 5, 1)
    global_var.teamb3ap = round(float(global_var.teamb3ap) / 5, 1)

    # The Formula
    global_var.teamb1ranking = int(10 - int(global_var.teamb1ranking))
    global_var.teamb2ranking = int(10 - int(global_var.teamb2ranking))
    global_var.teamb3ranking = int(10 - int(global_var.teamb3ranking))

    # /17
    global_var.teamb1highest = round(float(int(global_var.teamb1highest) / 17), 1)
    global_var.teamb2highest = round(float(int(global_var.teamb2highest) / 17), 1)
    global_var.teamb3highest = round(float(int(global_var.teamb3highest) / 17), 1)

    if int(global_var.teamb1ranking) <= 0:
        global_var.teamb1ranking = 0
    if int(global_var.teamb2ranking) <= 0:
        global_var.teamb2ranking = 0
    if int(global_var.teamb3ranking) <= 0:
        global_var.teamb3ranking = 0

    # Check
    print("Skill " + str(global_var.teamb1skillout) + " " + str(global_var.teamb2skillout) + " " + str(
        global_var.teamb3skillout))
    print("Season Wins " + str(global_var.teamb1wins) + " " + str(global_var.teamb2wins) + " " + str(
        global_var.teamb3wins))
    print("AP " + str(global_var.teamb1ap) + " " + str(global_var.teamb2ap) + " " + str(global_var.teamb3ap))
    print("Ranking " + str(global_var.teamb1ranking) + " " + str(global_var.teamb2ranking) + " " + str(
        global_var.teamb3ranking))
    print("Highest " + str(global_var.teamb1highest) + " " + str(global_var.teamb2highest) + " " + str(
        global_var.teamb3highest))

    df = pd.DataFrame.from_items([
        ('group', ['A', 'B', 'C', 'D']),
        ('Skill', [global_var.teamb1skillout, global_var.teamb2skillout, global_var.teamb3skillout, 0]),
        ('Season Wins', [global_var.teamb1wins, global_var.teamb2wins, global_var.teamb3wins, 0]),
        ('AP', [global_var.teamb1ap, global_var.teamb2ap, global_var.teamb3ap, 0]),  # /5
        ('Highest', [global_var.teamb1highest, global_var.teamb2highest, global_var.teamb3highest, 0]),  # /15
        ('Rankings', [global_var.teamb1ranking, global_var.teamb2ranking, global_var.teamb3ranking, 0])  # The Formula
    ])

    # 'Skill': [global_var.teamb1skillout, global_var.teamb2skillout, global_var.teamb3skillout,0],
    # 'Season Wins': [global_var.teamb1wins, global_var.teamb2wins, global_var.teamb3wins, 0],

    # ------- PART 1: Create background

    # number of variable

    categories = list(df)[1:]
    N = len(categories)

    # What will be the angle of each axis in the plot? (we divide the plot / number of variable)
    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]

    # Initialise the spider plot
    ax = plt.subplot(111, polar=True)

    # If you want the first axis to be on top:
    ax.set_theta_offset(pi / 2)
    ax.set_theta_direction(-1)

    # Draw one axe per variable + add labels labels yet
    plt.xticks(angles[:-1], categories)

    # Draw ylabels
    ax.set_rlabel_position(0)
    plt.yticks([3, 6, 9], ["3", "6", "9"], color="grey", size=7)
    plt.ylim(0, 12)

    # Ind1
    values = df.loc[0].drop('group').values.flatten().tolist()
    values += values[:1]
    ax.plot(angles, values, linewidth=1, linestyle='solid', label=str(global_var.teamb1))
    ax.fill(angles, values, 'b', alpha=0.1)

    # Ind2
    values = df.loc[1].drop('group').values.flatten().tolist()
    values += values[:1]
    ax.plot(angles, values, linewidth=1, linestyle='solid', label=str(global_var.teamb2))
    ax.fill(angles, values, 'r', alpha=0.1)

    # Ind3
    values = df.loc[2].drop('group').values.flatten().tolist()
    values += values[:1]
    ax.plot(angles, values, linewidth=1, linestyle='solid', label=str(global_var.teamb3))
    ax.fill(angles, values, 'r', alpha=0.1)

    # Add legend
    plt.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1))

    # plt.show("all")
    # plt.show("all")
    plt.show()
    plt.close()


# Start!


print(
    "[VEXDB Reader] By Team 35211C, Haorui Zhou, Yifei Ding \n Version 1.2 Update: 2018/4/25 21:11 \n Copyright: "
    "Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License \n Learn more about CC BY-NC-SA "
    "4.0: Choose '5' in Mode \n Contact Info: Discord Yingfeng#8524 \n")
time.sleep(0.5)
input("Press Any Key to Start!\n")
while True:
    mode = int(input(
        "Mode \n 1.Scan Team Matches \n 2.Excel Functions [Not Finished] \n 3.Search Team Season History \n 5.For "
        "Copyright License\n 6.Discord Link \n 8.Get Important Info For a Team \n 9.Change Log\n 0.Quit \n"))
    if mode == 1:
        print("Mode = Scan Team Matches")
        time.sleep(0.3)
        scanteammatches()
    elif mode == 2:
        print("Mode = Excels")
        # sleeptimer = float(input("Set Sleep Time\n"))
        print(
            "1.Scan Teams \n2.Scan Matches [Don't use this]\n3.Write Team Important Data\n4.Don't Ues This\n5.Can "
            "Specific Match [PreSet World Championship]\n6.Get We Need")
        time.sleep(0.3)
        excelmode = int(input())
        if excelmode == 1:
            print("Mode = Scan Teams and Write to Excel")
            time.sleep(0.3)
            excelscanteams()
        elif excelmode == 2:
            print("Mode = Write Team Matches [Don't use this]")
            time.sleep(0.3)
            excelteammatches()
        elif excelmode == 3:
            print("Mode = Write Team Important Data in Excel")
            time.sleep(0.3)
            excelgetalldata()
        elif excelmode == 4:
            print("Mode = Scan Bugged Team [It will crash]")
            time.sleep(0.3)
            excelgetallbugs()
        elif excelmode == 5:
            print("Mode = Scan World Championship")
            time.sleep(0.3)
            excelscanworld()
        elif excelmode == 6:
            print("Mode = Scan We Need")
            time.sleep(0.3)
            excelgetweneed()
    elif mode == 3:
        print("Mode = Search Team History : Current Season")
        time.sleep(0.3)
        searchteamcurrentseason()
    elif mode == 4:
        print("Bubble!")
        timeisout()
        answer()
    elif mode == 7:
        print("Mode = Empty")
        time.sleep(0.3)
        empty()
    elif mode == 8:
        print("Mode = Get Important Data")
        time.sleep(0.3)
        getalldata()
    elif mode == 0:
        print("Thanks for using it!")
        time.sleep(0.3)
        quit()
    else:
        print("Mode Unknown")
        time.sleep(1)
