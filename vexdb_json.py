"""
This is a rewrite of the old vexdb.io api reader, it read/ get json from vex.io and do something with them
Credit:Haorui Zhou (oraginal code), Yifei Ding(Rewrite)
license: CC By-NC-SA
contact: yifeiding@protonmail.com
"""
import json
from urllib.error import URLError
from urllib.request import urlopen

max_tries = 5  # increase this value when fail rate is very high


def get_json_url(url: str, safe=True, fail_counter=0):
    """
    This function get json directly from a url
    Change of max_tries will affect this function
    If max_tries is -1 the safe option is ignored
    :param url: The json url, should be http or https
    :param safe: if the safe option is off then it will not handle multiple tries
    :param fail_counter: you don't need to change this, it increase by 1 when a try failed
    :return: It return a dictionary of lists or tuples, not sure, dictionary of something, maybe json object??
    """
    if safe is False:
        return json.loads((urlopen(url)).read())
    if max_tries == -1:
        return json.loads((urlopen(url)).read())
    try:
        json_dict = json.loads((urlopen(url)).read())
    except URLError:
        if fail_counter < max_tries:
            get_json_url(url, safe, fail_counter + 1)
        else:
            raise ConnectionError("Multiple attempts to get data from vexdb.io failed, Abort")
    else:
        if json_dict["status"] != 1:
            raise TypeError("Unexpected Status")
        else:
            if json_dict["size"] == 5000:
                raise OverflowError("The Data size exceed 5000 item limit")
            if json_dict["size"] == 0:
                raise ValueError("doesn't contain any data")
            return json_dict


def get_json_direct(api_type: str, api_parameters: dict, safe=True):
    """
    it does additional data check on top of get_json, if you are okay with no data or exceed the max data then ignore it
    Change of max_tries will affect this function
    :param api_type: what is after get_ ?
    :param api_parameters: the parameters, should be the key:value of what you want to search
    :param safe: it go into get_json
    :return: It return a dictionary of lists or tuples, not sure, dictionary of something, maybe json object??
    """

    return get_json_url(url_gen(api_type, api_parameters), safe)


def url_gen(api_type: str, api_parameters: dict):
    """
    generate vexdb.io api url, nothing special
    :param api_type: what is after get_ ?
    :param api_parameters: the parameters, should be the key:value of what you want to search
    :return:a string that contain a url
    """
    _parameters = ""
    if api_type == "":
        raise GeneratorExit("Missing value for API type")

    if api_parameters:
        _keys = list(api_parameters.keys())
        _values = list(api_parameters.values())
        if len(_keys) >= 1:
            _parameters += "?" + _keys[0] + "=" + _values[0]
            if len(_keys) > 1:
                for x in range(1, len(_keys)):
                    if _values[x] == "":
                        raise GeneratorExit("The api_parameters must have non empty values")
                    _parameters += "&" + _keys[x] + "=" + _values[x]
    else:
        raise GeneratorExit("missing required information")
    return "http://api.vexdb.io/v1/get_" + api_type + _parameters


def filter_info(info: dict, *args: str):
    """
    filter the json from get_info or a similar structure dictionary and return a list
    :param info: this should be the dictionary from get_info or same structure
    :param args:the info key you want to contain, highly suggest only 1 arg
    :return:It return a list or 2d list that contains the value from info, all value in list are string
    """
    out = []
    if len(args) == 1:
        for item in info["result"]:
            for arg in args:
                out.append(item[arg])
        return out
    if len(args) > 1:
        for item in info["result"]:
            items = []
            for arg in args:
                items.append(str(item[arg]))
            out.append(items)
        return out


def check_info(api_type: str, info_type: str, api_parameter: str, safe=True):
    """
    Check if something is exit in vexdb.io. If you use it incorrectly it will return weird things for sure
    If something is wrong, it will raise a ValueError
    :param api_type: for example: teams
    :param info_type: what kind of data the parameter is?
    :param api_parameter: The thing you want to check if exit
    :param safe: directly go into get_json
    :return: It return a boolean, True means it exit (returned more than 0 item), False means it doesn't exit.
    """
    json_dict = get_json_url("http://api.vexdb.io/v1/get_" + api_type + "?" + info_type + "=" + api_parameter, safe)
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
