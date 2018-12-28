import vexdb_json, vexdb_excel


def getteam(sku, country):
    _json_dict = vexdb_json.get_json_direct("teams",
                                            {"sku": sku, "program": "VRC", "limit_number": "4999", "country": country})
    output = []
    for r in _json_dict["result"]:
        line = '{}: '.format(r["number"])
        output.append(line)
    return output


def main():
    a = vexdb_excel.WriteWorkbook()
    a.write_chart("test", [[{"test": (a.YELLOW_FILL, a.BOLD_BLUE_FONT)}]])
    a.save()


if __name__ == '__main__':
    main()
    # self.ranking_columns = "Team", "Wins", "Losses", "AP", "Ranking", "Highest", "Result"
    # self.matches_columns = "Sku", "Match", "Red1", "Red2", "Red3", "RedSit", "Blue1", "Blue2", "Blue3", "BlueSit", "RedSco", "BlueSco"
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
