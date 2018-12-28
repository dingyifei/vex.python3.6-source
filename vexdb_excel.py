"""
This is a rewrite of the vexdb excel visualizer module, it don't need internet connection if you use offline data
Credit:Haorui Zhou (oraginal code), Yifei Ding(Rewrite)
license: CC By-NC-SA
contact: yifeiding@protonmail.com
"""
import time
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors, Color, fonts
import vexdb_json


class WriteWorkbook:  # testing

    def __init__(self):

        # self.ranking_columns = "Team", "Wins", "Losses", "AP", "Ranking", "Highest", "Result"
        self.save_location = "./output.xlsx"
        # self.matches_columns = "Sku", "Match", "Red1", "Red2", "Red3", "RedSit", "Blue1", "Blue2", "Blue3", "BlueSit", "RedSco", "BlueSco"
        # self.sheet_names = "#Cover", "#Rankings", "#Important Data", "#For World", "#Bugged Teams"
        self.RED_FILL = PatternFill(patternType="solid", fgColor=colors.RED)
        self.BLUE_FILL = PatternFill(patternType="solid", fgColor=colors.BLUE)
        self.GREEN_FILL = PatternFill(patternType="solid", fgColor=colors.GREEN)  # replace Light Red
        self.YELLOW_FILL = PatternFill(patternType="solid", fgColor=colors.YELLOW)  # replace Light_Blue
        self.BLACK_FILL = PatternFill(patternType="solid", fgColor=colors.BLACK)
        self.BOLD_RED_FONT = Font("Calibri", size=11, color=colors.RED, bold=True)
        self.BOLD_BLUE_FONT = Font("Calibri", size=11, color=colors.BLUE, bold=True)
        self.BOLD_BLACK_FONT = Font("Calibri", size=11, color=colors.BLACK, bold=True)
        self.BOLD_WHITE_FONT = Font("Calibri", size=11, color=colors.WHITE, bold=True)
        self.book = openpyxl.Workbook()
        self.book.create_sheet("Cover")
        del self.book["Sheet"]  # I don't know how to solve this myth, it automatically generate sheets
        self.book["Cover"].cell(row=1, column=1).value = "Last Change:" + str(time.localtime())

    '''
    list of text - dictionary [key=the text, value=(font, fill)]
    '''

    @staticmethod
    def value_check(values: list):

        for x in values:
            if type(x) is tuple:
                if len(x) == 2:
                    if type(x[0]) != PatternFill:
                        raise ValueError("The first value in tuple should be PatternFill")
                    if type(x[1]) != Font:
                        raise ValueError("The second value in tuple should be Font")
                else:
                    raise ValueError("invalid tuple length")
            else:
                raise TypeError("The value should be tuple contain two value")

    def write_chart(self, sheet_name: str, text: dict, start_row=1, start_column=1):
        # TODO: The CODE INSIDE IS NOT WORKING
        if sheet_name not in self.book.get_sheet_names:
            self.book.create_sheet(sheet_name)
        active_sheet = self.book[sheet_name]
        keys: list = text.keys()
        values: list = text.values()
        self.value_check(values)

    # for x, y in enumerate(self.ranking_columns):  # Initialize Matches
    #     self.book[self.sheet_names[1]].cell(row=1, column=x + 1).value = y
    #     self.book[self.sheet_names[1]].cell(row=1, column=x + 1).font = self.BOLD_BLACK_FONT

    def save(self):
        self.book.save(self.save_location)


#  "Because of there are no data for these teams: 1119S, 7386A, 8000X, 8000Z, 19771B, 30638A, 36632A, "
#  "37073A, 60900A, 76921B, 99556A, 99691E, 99691H are not include in the sheet #Important Data")


def team_list():  # For testing

    # print(vexdb_json("teams", {"grade": "High%20School"}, ["number"]))
    print(vexdb_json.get_json_direct("matches", {"season": "Starstruck", "team": "8667A"}, ["sku"]))


def getteam(sku, country):
    # TODO: fix after finish readconfig
    _json_dict = vexdb_json.get_json_direct("teams",
                                            {"sku": sku, "program": "VRC", "limit_number": "4999", "country": country})
    output = []
    for r in _json_dict["result"]:
        line = '{}: '.format(r["number"])
        output.append(line)
    return output


def main():
    a = WriteWorkbook()
    a.save()


if __name__ == '__main__':
    main()

    #
    # def rankings_excel(self):
    #     rankings_columns = ("Team", "Wins", "Losses", "AP", "Ranking", "Highest", "Result")
    #     for x, y in enumerate(rankings_columns):  # Initialize Matches
    #         book[sheet_names[1]].cell(row=1, column=x + 1).value = y
    #         book[sheet_names[1]].cell(row=1, column=x + 1).font = ExcelStyle.BOLD_BLACK_FONT
    #
    #     # for row, team in enumerate(rankings_scan(teams, season)):
    #     #     for column, value in enumerate(team):
    #     #         book[sheet_names[1]].cell(row=row + start_row, column=6 + start_column).value = value
    #     #     if int(team[1]) > int(team[2]):
    #     #         book[sheet_names[1]].cell(row=row + start_row + 1, column=6 + start_column).value = "Positive"
    #     #         book[sheet_names[1]].cell(row=row + start_row + 1, column=6 + start_column).fill = ExcelStyle.RED_FILL
    #     #     elif int(team[1]) < int(team[2]):
    #     #         book[sheet_names[1]].cell(row=row + start_row + 1, column=6 + start_column).value = "Negative"
    #     #         book[sheet_names[1]].cell(row=row + start_row + 1, column=6 + start_column).fill = ExcelStyle.BLUE_FILL
    #     #     elif int(team[1]) == int(team[2]):
    #     #         book[sheet_names[1]].cell(row=row + start_row + 1, column=6 + start_column).value = "Equal"
    #     #         book[sheet_names[1]].cell(row=row + start_row, column=6 + start_column).fill = ExcelStyle.BLACK_FILL
    #     #     else:
    #     #         book[sheet_names[1]].cell(row=row + start_row + 1, column=6 + start_column).value = "Error"
    #     #         book[sheet_names[1]].cell(row=row + start_row + 1, column=6 + start_column).fill = ExcelStyle.GREEN_FILL
    #     #     book[sheet_names[1]].cell(row=row + 1, column=6 + start_column).font = ExcelStyle.BOLD_WHITE_FONT
    #
    # def matches_excel(self):
    #     matches_columns = (
    #         "Sku", "Match", "Red1", "Red2", "Red3", "RedSit", "Blue1", "Blue2", "Blue3", "BlueSit", "RedSco", "BlueSco"
    #     )
    #     for x, y in enumerate(matches_columns):  # Initialize Matches
    #         book[sheet_names[1]].cell(row=1, column=x + 1).value = y
    #         book[sheet_names[1]].cell(row=1, column=x + 1).font = ExcelStyle.BOLD_BLACK_FONT
    #
    #     # for row, match in enumerate(matches_scan(team, season)):
    #     #     for column, value in enumerate(match):
    #     #         book[sheet_names[1]].cell(row=start_row + row, column=start_column + column)
    #     #     if int(match[10]) > int(match[11]):
    #     #         book[sheet_names[1]].cell(row=row + 1, column=start_column + 14).value = "Red"
    #     #         book[sheet_names[1]].cell(row=row + 1, column=start_column + 14).fill = ExcelStyle.RED_FILL
    #     #     elif int(match[10]) < int(match[11]):
    #     #         book[sheet_names[1]].cell(row=row + 1, column=start_column + 14).value = "Blue"
    #     #         book[sheet_names[1]].cell(row=row + 1, column=start_column + 14).fill = ExcelStyle.BLUE_FILL
    #     #     if int(match[10]) + 20 < int(match[11]):
    #     #         book[sheet_names[1]].cell(row=row + 1, column=start_column + 14).value = "Blue Easy"
    #     #         book[sheet_names[1]].cell(row=row + 1, column=start_column + 14).fill = ExcelStyle.GREEN_FILL
    #     #     elif int(match[10]) - 20 > int(match[11]):
    #     #         book[sheet_names[1]].cell(row=row + 1, column=start_column + 14).value = "Red Easy"
    #     #         book[sheet_names[1]].cell(row=row + 1, column=start_column + 14).fill = ExcelStyle.YELLOW_FILL
    #     #
    #     #     book[sheet_names[1]].cell(row=row + 1, column=start_column + 14).font = ExcelStyle.BOLD_WHITE_FONT
    #     #
    #     #     if match[2] == team or match[3] == team or match[4] == team:
    #     #         if int(match[10]) > int(match[11]):
    #     #             book[sheet_names[1]].cell(row=row + 1, column=start_column + 13).value = "Win"
    #     #             book[sheet_names[1]].cell(row=row + 1, column=start_column + 13).font = ExcelStyle.BOLD_BLACK_FONT
    #     #         else:
    #     #             book[sheet_names[1]].cell(row=row + 1, column=start_column + 13).value = "Lose"
    #     #             book[sheet_names[1]].cell(row=row + 1, column=start_column + 13).font = ExcelStyle.BOLD_WHITE_FONT
    #     #             book[sheet_names[1]].cell(row=row + 1, column=start_column + 13).fill = ExcelStyle.BLACK_FILL
    #     #     elif match[6] == team or match[7] == team or match[8] == team:
    #     #         if int(match[10]) < int(match[11]):
    #     #             book[sheet_names[1]].cell(row=row + 1, column=start_column + 13).value = "Win"
    #     #             book[sheet_names[1]].cell(row=row + 1, column=start_column + 13).font = ExcelStyle.BOLD_BLACK_FONT
    #     #         else:
    #     #             book[sheet_names[1]].cell(row=row + 1, column=start_column + 13).value = "Lose"
    #     #             book[sheet_names[1]].cell(row=row + 1, column=start_column + 13).font = ExcelStyle.BOLD_WHITE_FONT
    #     #             book[sheet_names[1]].cell(row=row + 1, column=start_column + 13).fill = ExcelStyle.BLACK_FILL
