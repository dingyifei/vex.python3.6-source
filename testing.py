import sys
import json
import vexdb_json
import unittest

a = open("testing.json")

a_json = json.loads(a.read())
print(a_json)
print(vexdb_json.filter_info(a_json, "number", "program"))



class IntegerArithmeticTestCase(unittest.TestCase):
    def testAdd(self):  # test method names begin with 'test'
        self.assertEqual((1 + 2), 3)
        self.assertEqual(0 + 1, 1)

    def testMultiply(self):
        self.assertEqual((0 * 10), 0)
        self.assertEqual((5 * 8), 40)


if __name__ == '__main__':
    unittest.main()
