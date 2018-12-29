import sys
import json
import vexdb_json
import unittest
import vexdb_excel
import openpyxl

# TOO MUCH WORK< IGNORE THIS
class IntegerArithmeticTestCase(unittest.TestCase):
    # a = open("testing.json")
    #
    # a_json = json.loads(a.read())
    # print(a_json)
    # print(vexdb_json.filter_info(a_json, "number", "program"))

    def testFormater(self): # I have no idea why it is not working
        a = vexdb_excel.WriteWorkbook()
        print(a.list_formater([["testing"]]))
        self.assertEqual(a.list_formater([["testing"]]),
                         [[["testing", a.WHITE_FILL, a.BOLD_BLACK_FONT]]]
                         )



if __name__ == '__main__':
    unittest.main()


    # def team_list():  # For testing
    #
    #     # print(vexdb_json("teams", {"grade": "High%20School"}, ["number"]))
    #     print(vexdb_json.get_json_direct("matches", {"season": "Starstruck", "team": "8667A"}, ["sku"]))
