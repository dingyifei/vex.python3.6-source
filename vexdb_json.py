"""
TODO: Reserved docstring space
"""
import json
import threading
from urllib.error import URLError
from urllib.request import urlopen

max_tries = 5  # increase this value when fail rate is very high


def get_info(api_type: str, api_parameters: dict, safe=True):
    """
    It function accept a string "api_type" and a dictionary "api_parameters", the "api_type" should be
    one from _API_TYPE The dictionary's key are the _parameters from vexdb.io/the_data and the value should
    also follow it.
    All ValueError are critical
    :param api_type: what is after get_ ?
    :param api_parameters: the parameters, should be the key:value of what you want to search
    :param safe: it go into get_json
    :return: It return a dictionary of lists or tuples, not sure, dictionary of something
    """
    # TODO(Yifei): Multi thread,

    _parameters = ""

    if api_type == "":
        raise ValueError("Missing value for API type")

    if api_parameters:
        _keys = list(api_parameters.keys())
        _values = list(api_parameters.values())
        if len(_keys) >= 1:
            _parameters += "?" + _keys[0] + "=" + _values[0]
            if len(_keys) > 1:
                for x in range(1, len(_keys)):
                    _parameters += "&" + _keys[x] + "=" + _values[x]
    else:
        raise ValueError("missing required information")
        exit(1)

    if api_type != "":
        json_dict = get_json("http://api.vexdb.io/v1/get_" + api_type + _parameters, safe)
        if json_dict["status"] != 1:
            raise TypeError("Unexpected Status")
        else:
            if json_dict["size"] == 5000:
                raise OverflowError("The Data size exceed 5000 item limit")
            return json_dict
    else:
        raise ValueError("missing required information")
        exit(1)


def get_json(url: str, safe=True, fail_counter=0):
    if max_tries == -1:  # force to ignore safe
        return json.loads((urlopen(url)).read())
    try:
        out = json.loads((urlopen(url)).read())
    except URLError:
        if fail_counter < max_tries:
            get_json(url, safe, fail_counter + 1)
        else:
            raise ConnectionError("Multiple attempts to get data from vexdb.io failed, Abort")
    else:
        return out


def filter_info(info: dict):

    print("the return data thing")


def check_info(api_type: str, info_type: str, api_parameter: str, safe=True):
    """
    Check if something is exit in vexdb.io. If you use it incorrectly it will return weird things for sure
    :param api_type: for example: teams
    :param info_type: what kind of data the parameter is?
    :param api_parameter: The thing you want to check if exit
    :param safe: directly go into get_json
    :return: It return a boolean, True means it exit (returned more than 0 item), False means it doesn't exit.
    :return: If something is wrong, it will raise a ValueError
    """
    json_dict = get_json("http://api.vexdb.io/v1/get_" + api_type + "?" + info_type + "=" + api_parameter, safe)
    try:
        if json_dict["size"] > 0:
            return True
        else:
            return False
    except KeyError:  # at this point things 100% go wrong so time for the returns
        try:
            raise ValueError(json_dict["error_text"])
        except KeyError:
            raise ValueError("Unexpected Error")


def main():
    """
    you should not use it directly
    """
    print(check_info("teams", "team", "2915A"))


if __name__ == '__main__':
    main()
