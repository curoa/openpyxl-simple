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
