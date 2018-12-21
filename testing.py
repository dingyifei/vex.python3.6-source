import sys
import json
import vexdb_json
a = open("testing.json")

a_json = json.loads(a.read())
print(a_json)
print(vexdb_json.filter_info(a_json, "number", "program"))