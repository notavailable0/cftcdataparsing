from parsing import parsing
from excel_managment import i_fucked_excel
import json


def main(url_list):
    out = []
    for i in url_list:
        r = parsing(i)
        for a in r:
            out.append(a)

    return out




if __name__ == "__main__":
    names_raw = []
    names = []
    outs = []
    instert_to_excel = []
    url_list = ['https://www.cftc.gov/dea/futures/deanybtsf.htm', 'https://www.cftc.gov/dea/futures/deaiceusf.htm', 'https://www.cftc.gov/dea/futures/deaifedsf.htm', 'https://www.cftc.gov/dea/futures/deacbtsf.htm', 'https://www.cftc.gov/dea/futures/deacmesf.htm', 'https://www.cftc.gov/dea/futures/deacboesf.htm', 'https://www.cftc.gov/dea/futures/deaccxsf.htm', 'https://www.cftc.gov/dea/futures/deakcbtsf.htm', 'https://www.cftc.gov/dea/futures/deamgesf.htm', 'https://www.cftc.gov/dea/futures/deacmxsf.htm', 'https://www.cftc.gov/dea/futures/deanymesf.htm', 'https://www.cftc.gov/dea/futures/deanylsf.htm', 'https://www.cftc.gov/dea/futures/deanypcsf.htm', 'https://www.cftc.gov/dea/futures/deanodxsf.htm', 'https://www.cftc.gov/dea/futures/deadumxsf.htm', 'https://www.cftc.gov/dea/futures/deapbotsf.htm', 'https://www.cftc.gov/dea/futures/deaerissf.htm']

    out = main(url_list)

    with open('names.txt', 'r') as nm:
        names_raw = nm.readlines()

    for i in names_raw:
        names.append(i.strip('\n'))

    with open('output.txt', 'w') as d1:
        for i in out:
            instert_to_excel.append(i)
            if i['name'] in names:
                outs.append(i)
                d1.write(json.dumps(i)+'\n')

    with open('data_all.txt', 'w') as da:
        for i in out:
            print(i['name'])
            da.write(json.dumps(i)+'\n')

    print('successfully parsed all the values')
    print(out)

    i_fucked_excel(outs)

