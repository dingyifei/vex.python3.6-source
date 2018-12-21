"""
TODO: Reserved docstring space
"""
import json
from urllib.request import urlopen

log_output = print


def get_info(api_type: str, api_parameters: dict, return_data=None):
    """
    It function accept a string "api_type" and a dictionary "api_parameters", the "api_type" should be
    one from _API_TYPE The dictionary's key are the _parameters from vexdb.io/the_data and the value should
    also follow it.
    :param api_type: what is after get_ ?
    :param api_parameters: the parameters, should be the key:value of what you want to search
    :param return_data: it is a simple filter that make return data be a list of weird returns, not fully functional
    :return: It return a dictionary or a list, depend on what you do TODO: Make sure it is a dictionary
    """
    # TODO(Yifei): Multi thread, timeout retry, what will happen when no interenet?

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
                raise TypeError("Unexpected Status")
            else:
                if json_dict["size"] == 5000:
                    raise OverflowError("The Data size exceed 5000 item limit")
                else:
                    if return_data[0] == "full":
                        output: dict = json_dict
                    if return_data[0] != "full":
                        output: list = []
                        for x in json_dict["result"]:
                            for y in return_data:
                                output.append(x[y])
                return output


def get_info_safe():
    # TODO: handle fail and an option to retry
    print("Not working")


def get_info_multi_thread():
    # TODO: get info like a cool person with syncornized threads, should able to get more than 5000 values?
    print("not working")


def get_info_multi_thread_safe():
    # TODO: Safe version of multi thread, otherwise it will just throw errors or something IDK
    print("not working")


def check_info(api_type: str, info_type: str, api_parameter: str):
    """
    Check if something is exit in vexdb.io. If you use it incorrectly it will return weird things for sure
    :param api_type: for example: teams
    :param info_type: what kind of data the parameter is?
    :param api_parameter: The thing you want to check if exit
    :return: It return a boolean, True means it exit (returned more than 0 item), False means it doesn't exit.
    :return: If something is wrong, it will raise a ValueError

    """
    json_dict = json.loads(
        (urlopen("http://api.vexdb.io/v1/get_" + api_type + "?" + info_type + "=" + api_parameter)).read())
    try:
        if json_dict["size"] > 0:
            return True
        else:
            return False
    except KeyError:
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
