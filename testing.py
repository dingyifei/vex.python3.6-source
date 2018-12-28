import sys
import json
import vexdb_json
import unittest


# TOO MUCH WORK< IGNORE THIS
class IntegerArithmeticTestCase(unittest.TestCase):
    a = open("testing.json")

    a_json = json.loads(a.read())
    print(a_json)
    print(vexdb_json.filter_info(a_json, "number", "program"))

    def testAdd(self):  # test method names begin with 'test'
        self.assertEqual((1 + 2), 3)
        self.assertEqual(0 + 1, 1)

    def testMultiply(self):
        self.assertEqual((0 * 10), 0)
        self.assertEqual((5 * 8), 40)


if __name__ == '__main__':
    unittest.main()


    def team_list():  # For testing

        # print(vexdb_json("teams", {"grade": "High%20School"}, ["number"]))
        print(vexdb_json.get_json_direct("matches", {"season": "Starstruck", "team": "8667A"}, ["sku"]))
