from functools import wraps
from inspect import getfullargspec
from pathlib import Path

def to_path(fpath):
    if type(fpath) == str:
        fpath = Path(fpath)
    return fpath

def deco_fname_check(ftype):
    def _deco_fname_check(f):
        argspec = getfullargspec(f)
        argument_name = "fpath"
        argument_index = argspec.args.index(argument_name)
        @wraps(f)
        def wrapper(*args, **kwargs):
            if argument_index < len(args):
                fpath = to_path(args[argument_index])
            else:
                fpath = to_path(kwargs[argument_name])
            if fpath.suffix[1:] != ftype:
                msg = f"file name maybe wrong: expected ends with {ftype} but get {fpath}"
                print("\033[31m" + msg + "\033[0m") # ]] fix vim indent
            # do something with value
            return f(*args, **kwargs)
        return wrapper
    return _deco_fname_check

def guess_utf_encoding(fpath):
    with open(fpath, "rb") as f:
        beginning = f.read(4)
        #NOTE The order of these if-statements is important. otherwise UTF32 LE may be detected as UTF16 LE as well.
        #NOTE delete -le -be to skip BOM.  ref. https://docs.python.org/ja/3.7/howto/unicode.html#reading-and-writing-unicode-data
        if beginning == b"\x00\x00\xfe\xff":
            return "utf-32"
            return "utf-32-be"
        elif beginning == b"\xff\xfe\x00\x00":
            return "utf-32"
            return "utf-32-le"
        elif beginning[0:3] == b"\xef\xbb\xbf":
            return "utf-8-sig"
        elif beginning[0:2] == b"\xff\xfe":
            return "utf-16"
            return "utf-16-le"
        elif beginning[0:2] == b"\xfe\xff":
            return "utf-16"
            return "utf-16-be"
        else:
            #NOTE handle unknown as utf-8
            return "utf-8"

def tie_key_value(header, value_list):
    data = []
    for row in value_list:
        new = {}
        for k, v in zip(header, row):
            new[k] = v
        data.append(new)
    return data

def collect_header(data):
    s = set() # use set to check already registered
    header = [] # use list to save order
    for row in data:
        l = list(row.keys())
        for column_name in l:
            if column_name in s:
                continue
            s.add(column_name)
            header.append(column_name)
    return header
