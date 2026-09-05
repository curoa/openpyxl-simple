from functools import wraps
from inspect import getfullargspec, signature, unwrap
from pathlib import Path

def deco_to_path(fn_or_arg_name=None, arg_name="fpath"):
    target_arg_name = fn_or_arg_name if isinstance(fn_or_arg_name, str) else arg_name

    def _deco_to_path(func):
        sig = signature(func)

        @wraps(func)
        def wrapper(*args, **kwargs):
            bound = sig.bind(*args, **kwargs)
            bound.apply_defaults()
            val = bound.arguments.get(target_arg_name)
            if isinstance(val, str):
                bound.arguments[target_arg_name] = Path(val)
            return func(*bound.args, **bound.kwargs)

        return wrapper

    return _deco_to_path(fn_or_arg_name) if callable(fn_or_arg_name) else _deco_to_path

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
