import requests
import re

##def find_sub(s, f): return s[]

def parsing(link):
    raw = requests.get(link).text
    #print(raw)

    row_splitted = raw.split('\n')

    data_1 = {
    "name" : "",
    "code" : "",
    "date" : "",
    "open_interest_all" : '',
    "non_commercial_long" : '',
    "non_commercial_short" : '' }

    json_out = []

    for i in row_splitted:
        if 'Code' in i:
            print(i)
            try:
                data_1["name"] = i.split('   ')[0]
                data_1["code"] = "Code" + i.split('Code')[1].strip()
                ## print(data_1['name'], end = '\n')

                index_i = row_splitted.index(i)
                # print(index_i)

                date = row_splitted[index_i + 1].split('OF ')[1].split('  ')[0]
                data_1["date"] = date
                # print(date)

                val_list = row_splitted[index_i+9].split('  ')
                for i in val_list:
                    i = i.strip()
                    if len(i) == 0:
                        val_list.remove(i)

                # print(val_list)

                open_interest = row_splitted[index_i + 7].split(':')[1].strip()
                # print(open_interest)
                data_1["open_interest_all"] = row_splitted[index_i + 7].split(':')[1].strip()

                data_1["non_commercial_long"] = val_list[3].strip()
                data_1["non_commercial_short"] = val_list[4].strip()
                json_out.append(dict(data_1))
            except Exception as e:
                print(e)

    for i in json_out:
        print(i)

    return json_out