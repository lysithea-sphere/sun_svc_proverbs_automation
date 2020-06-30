def eng_str_break(str):     #function takes a string and returns a list of finalized list
    eng_limit = 108  # limit per slide
    empty_list = []
    if len(str) <= eng_limit:
        empty_list.append(str)
    else:
        while len(str) > eng_limit:
            comma_pos = str.find(',', 75, 103)  # it returns -1 integer when false
            fstop_pos = str.find('.', 75, 103)  # it returns -1 integer when false
            if comma_pos == -1 and fstop_pos == -1:  # comma and fs NOT Present
                space_pos = str.find(' ', 85, 103)  # might need to increase range by changing lower limit, in case last word of proverb is too long
                p1 = str[:space_pos + 1] + '[...]'  # check if space_pos needs to +1
                empty_list.append(p1)
                str = str[space_pos + 1:]
                continue
            elif comma_pos != -1 and fstop_pos == -1:  # comma present within range, no fs found
                p1 = str[:comma_pos + 1] + '[...]'
                empty_list.append(p1)
                str = str[comma_pos + 1:]
                continue
            elif comma_pos == -1 and fstop_pos != -1:  # comma NOT present within range, fs PRESENT
                p1 = str[: fstop_pos + 1] + '[...]'
                empty_list.append(p1)
                str = str[fstop_pos + 1:]
                continue
            else:
                p1 = str[: fstop_pos + 1] + '[...]'  # comma and fstop BOTH Present, but Fstop is chosen
                empty_list.append(p1)
                str = str[fstop_pos + 1:]
                continue
        empty_list.append(str)
    return empty_list


def chi_str_break(str):     #function takes a string and returns a list of finalized list
    chi_limit = 63  # 21 characters per line; limit per slide (chinese auto monospace?)
    empty_list = []
    if len(str) <= chi_limit:
        empty_list.append(str)
    else:
        while len(str) >= chi_limit:
            comma_pos = str.find('，', 44, 60)  # it returns -1 integer when false，
            fstop_pos = str.find('。', 44, 60)  # it returns -1 integer when false
            if comma_pos == -1 and fstop_pos == -1:  # comma and fs NOT Present
                p1 = str[:60] + '[...]'  # check if space_pos needs to +1
                empty_list.append(p1)
                str = str[60:]
                continue
            elif comma_pos != -1 and fstop_pos == -1:  # comma present within range, no fs found
                p1 = str[:comma_pos + 1] + '[...]'
                empty_list.append(p1)
                str = str[comma_pos + 1:]
                continue
            elif comma_pos == -1 and fstop_pos != -1:  # comma NOT present within range, fs PRESENT
                p1 = str[: fstop_pos + 1] + '[...]'
                empty_list.append(p1)
                str = str[fstop_pos + 1:]
                continue
            else:
                p1 = str[: fstop_pos + 1] + '[...]'  # comma and fstop BOTH Present, but Fstop is chosen
                empty_list.append(p1)
                str = str[fstop_pos + 1:]
                continue
        empty_list.append(str)
    return empty_list

def list_string_remover(list, string, newstring): # this function takes the list and replaces the string for each of
    # the item in the list, removes '-' and strips the spaces at the start and end
    list = [i.replace(string, newstring, 1) for i in list]
    list = [i.replace('-', '', 1) for i in list]
    list = [i.lstrip() for i in list]
    list = [i.rstrip() for i in list]
    return list