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
