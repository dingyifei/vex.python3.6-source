import json
from urllib.request import urlopen

def vexdb_json(api_type: str, api_parameters: dict, return_data=None):
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
            json_dict = json.loads((urlopen("http://api.vexdb.io/v1/get_" + api_type + _parameters)).read())
            if json_dict["status"] == 0:
                raise (IOError)  # TODO: a exception
            else:
                if json_dict["size"] == 5000:
                    raise (IOError)  # TODO: Another exception or use some trick to prevent 5000 limit
                else:
                    if return_data[0] == "full":
                        output: dict = json_dict
                    if return_data[0] != "full":
                        output: list = []
                        for x in json_dict["result"]:
                            for y in return_data:
                                output.append(x[y])
                return output


def main():
    print("Helloworld")


if __name__ == '__main__':
    main()
