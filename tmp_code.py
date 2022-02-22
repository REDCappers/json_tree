def json_tree_success(data, tab=0):
    indent = tab * '\t'
    if type(data) == dict:
        for k in data.keys():
            print('\n', indent, k, end='')
            json_tree(data[k], tab + 1)
    elif type(data) == list and len(data) > 0:
        for d in data:
            json_tree(d, tab)  # 改行なし
    else:
        print(' :', data, end='')
    return


def json_tree_success_2(data, tab=0):
    global string
    indent = tab * '\t'
    if type(data) == dict:
        for k in data.keys():
            # print('\n', indent, k, end='')
            string += '\n' + indent + k
            json_tree(data[k], tab + 1)
    elif type(data) == list and len(data) > 0:
        for d in data:
            # json_tree(data[0], tab + 1)
            json_tree(d, tab)  # 改行なし
    else:
        string += ' :' + str(data)
    return