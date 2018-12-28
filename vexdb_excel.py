"""
This is a rewrite of the vexdb excel visualizer module, it don't need internet connection if you use offline data
Credit:Haorui Zhou (oraginal code), Yifei Ding(Rewrite)
license: CC By-NC-SA
contact: yifeiding@protonmail.com
"""
import time

import openpyxl
from openpyxl.styles import PatternFill, Font, colors


class WriteWorkbook:  # testing
    """
    if you see this, you better check each function's docstring because I don't know how to explain this
    """
    def __init__(self):

        # self.ranking_columns = "Team", "Wins", "Losses", "AP", "Ranking", "Highest", "Result"
        # self.matches_columns = "Sku", "Match", "Red1", "Red2", "Red3", "RedSit", "Blue1", "Blue2", "Blue3", "BlueSit", "RedSco", "BlueSco"

        self.save_location = "./output.xlsx"
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

    @staticmethod
    def value_check(values: list):
        """
        it check the 2d list is the correct format
        :param values: the 2d list, it will make sure the 2d list is correct format
        """
        try:
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
        except:
            raise ValueError("Something big about the value is wrong")

    def write_chart(self, sheet: str, text: list, start_row=1, start_column=1):
        """
        This function write chart, YOU MUST USE THE CORRECT FORMAT 2D LIST!!!!
        :param sheet: if the sheet does not exit, it will be created automatically
        :param text: a 2d list, for example: [[{"test": (a.YELLOW_FILL, a.BOLD_BLUE_FONT)}]]
        :param start_row: the start row, should be the top row
        :param start_column: the start column, should be the left column
        """
        # TODO: The CODE INSIDE IS NOT WORKING
        if sheet not in self.book.sheetnames:
            self.book.create_sheet(sheet)
        active_sheet = self.book[sheet]
        for row, a in enumerate(text):
            for column, b in enumerate(a):
                self.value_check(b.values())
                active_sheet.cell(row=start_row + row, column=start_column + column).value = list(b.keys())[0]
                active_sheet.cell(row=start_row + row, column=start_column + column).fill = list(b.values())[0][0]
                active_sheet.cell(row=start_row + row, column=start_column + column).font = list(b.values())[0][1]

    def save(self):
        """
        it save the file
        """
        self.book.save(self.save_location)


#  "Because of there are no data for these teams: 1119S, 7386A, 8000X, 8000Z, 19771B, 30638A, 36632A, "
#  "37073A, 60900A, 76921B, 99556A, 99691E, 99691H are not include in the sheet #Important Data")


def main():
    """
    usually this main doesn't do anything
    """
    a = WriteWorkbook()
    a.write_chart("test", [[{"test": (a.YELLOW_FILL, a.BOLD_BLUE_FONT)}]])
    a.save()


if __name__ == '__main__':
    main()
