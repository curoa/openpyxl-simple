from functools import wraps
from inspect import getfullargspec, unwrap
from pathlib import Path

def deco_to_path(fn_or_arg_name=None, arg_name="fpath"):
    if callable(fn_or_arg_name):
        func = fn_or_arg_name
        target_arg_name = "fpath"
        argspec = getfullargspec(unwrap(func))
        arg_index = argspec.args.index(target_arg_name)

        @wraps(func)
        def wrapper(*args, **kwargs):
            if arg_index < len(args):
                val = args[arg_index]
                if isinstance(val, str):
                    args = list(args)
                    args[arg_index] = Path(val)
                    args = tuple(args)
            elif target_arg_name in kwargs:
                val = kwargs[target_arg_name]
                if isinstance(val, str):
                    kwargs[target_arg_name] = Path(val)
            return func(*args, **kwargs)

        return wrapper

    target_arg_name = fn_or_arg_name if isinstance(fn_or_arg_name, str) else arg_name

    def _deco_to_path(func):
        argspec = getfullargspec(unwrap(func))
        arg_index = argspec.args.index(target_arg_name)

        @wraps(func)
        def wrapper(*args, **kwargs):
            if arg_index < len(args):
                val = args[arg_index]
                if isinstance(val, str):
                    args = list(args)
                    args[arg_index] = Path(val)
                    args = tuple(args)
            elif target_arg_name in kwargs:
                val = kwargs[target_arg_name]
                if isinstance(val, str):
                    kwargs[target_arg_name] = Path(val)
            return func(*args, **kwargs)

        return wrapper

    return _deco_to_path

def deco_fname_check(ftype, arg_name="fpath"):
    def _deco_fname_check(f):
        argspec = getfullargspec(unwrap(f))
        argument_name = arg_name
        argument_index = argspec.args.index(argument_name)

        @wraps(f)
        def wrapper(*args, **kwargs):
            if argument_index < len(args):
                fpath = args[argument_index]
            else:
                fpath = kwargs[argument_name]
            if isinstance(fpath, str):
                fpath = Path(fpath)
            if fpath.suffix[1:] != ftype:
                msg = f"file name maybe wrong: expected ends with {ftype} but get {fpath}"
                print("\033[31m" + msg + "\033[0m")
            return f(*args, **kwargs)

        return wrapper

    return _deco_fname_check
