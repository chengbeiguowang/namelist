import re


def is_float(string):
    pattern = r'^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?$'
    if re.match(pattern, string):
        return True
    else:
        return False
