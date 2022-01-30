import requests
import re
from excel_managment import i_fucked_excel

names = {
    'S&P 500': 'S&P 500 Consolidated - CHICAGO MERCANTILE EXCHANGE',
    'S&P 500 Mini': 'E-MINI S&P 500 STOCK INDEX - CHICAGO MERCANTILE EXCHANGE',
    'Dow Jones': 'DJIA Consolidated - CHICAGO BOARD OF TRADE',
    'Dow Jones Mini': 'DOW JONES INDUSTRIAL AVG- x $5 - CHICAGO BOARD OF TRADE',
    'NASDAQ': 'NASDAQ-100 Consolidated - CHICAGO MERCANTILE EXCHANGE',
    'NASDAQ Mini': 'NASDAQ-100 STOCK INDEX (MINI) - CHICAGO MERCANTILE EXCHANGE',
    'RUSSEL 2000': 'E-MINI RUSSELL 2000 INDEX - CHICAGO MERCANTILE EXCHANGE',
    'NIKKEI': 'NIKKEI STOCK AVERAGE - CHICAGO MERCANTILE EXCHANGE',
    'VIX': 'VIX FUTURES - CBOE FUTURES EXCHANGE',
    '30 Year Bond': 'U.S. TREASURY BONDS - CHICAGO BOARD OF TRADE',
    '2 Year Note': '2-YEAR U.S. TREASURY NOTES - CHICAGO BOARD OF TRADE',
    '5 Year Note': '5-YEAR U.S. TREASURY NOTES - CHICAGO BOARD OF TRADE',
    '10 Year Note': '10-YEAR U.S. TREASURY NOTES - CHICAGO BOARD OF TRADE',
    'Federal Funds': '30-DAY FEDERAL FUNDS - CHICAGO BOARD OF TRADE',
    'Ultra U.S. T-Bonds': 'ULTRA U.S. TREASURY BONDS - CHICAGO BOARD OF TRADE',
    'Ultra 10-year U.S. T-Notes': 'ULTRA 10-YEAR U.S. T-NOTES - CHICAGO BOARD OF TRADE',
    'BitCoin': 'BITCOIN - CHICAGO MERCANTILE EXCHANGE',
    'Euro': 'EURO FX - CHICAGO MERCANTILE EXCHANGE',
    'Australischer Dollar': 'AUSTRALIAN DOLLAR - CHICAGO MERCANTILE EXCHANGE',
    'Brasilianischer Real': 'BRAZILIAN REAL - CHICAGO MERCANTILE EXCHANGE',
    'Britisches Pfund': 'BRITISH POUND STERLING - CHICAGO MERCANTILE EXCHANGE',
    'Japanischer Yen': 'JAPANESE YEN - CHICAGO MERCANTILE EXCHANGE',
    'Kanadischer Dollar': 'CANADIAN DOLLAR - CHICAGO MERCANTILE EXCHANGE',
    'Mexikanischer Peso': 'MEXICAN PESO - CHICAGO MERCANTILE EXCHANGE',
    'Neuseeland Dollar': 'NEW ZEALAND DOLLAR - CHICAGO MERCANTILE EXCHANGE',
    'Russischer Rubel': 'RUSSIAN RUBLE - CHICAGO MERCANTILE EXCHANGE',
    'Schweizer Franken': 'SWISS FRANC - CHICAGO MERCANTILE EXCHANGE',
    '3 Monate Euro Dollar': '3-MONTH EURODOLLARS - CHICAGO MERCANTILE EXCHANGE',
    'U.S. Dollar Index': 'U.S. DOLLAR INDEX - ICE FUTURES U.S.',
    'Gold': 'GOLD - COMMODITY EXCHANGE INC.',
    'Silber': 'SILVER - COMMODITY EXCHANGE INC.',
    'Platin': 'PLATINUM - NEW YORK MERCANTILE EXCHANGE',
    'Palladium': 'PALLADIUM - NEW YORK MERCANTILE EXCHANGE',
    'Kupfer': 'COPPER-GRADE #1 - COMMODITY EXCHANGE INC.',
    'Aluminium': 'ALUMINUM MW US TR PLATTS - COMMODITY EXCHANGE INC.',
    'Stahl': '',
    'Kohle': 'COAL (API 2) CIF ARA - NEW YORK MERCANTILE EXCHANGE',
    'Rohöl WTI': 'CRUDE OIL, LIGHT SWEET - NEW YORK MERCANTILE EXCHANGE',
    'Benzin': 'GASOLINE BLENDSTOCK (RBOB) - NEW YORK MERCANTILE EXCHANGE',
    'Heizöl': '#2 HEATING OIL- NY HARBOR-ULSD - NEW YORK MERCANTILE EXCHANGE',
    'Erdgas': 'NATURAL GAS - NEW YORK MERCANTILE EXCHANGE',
    'Ethanol': '',
    'Chicago Ethanol': 'CHICAGO ETHANOL - NEW YORK MERCANTILE EXCHANGE',
    'Mais': 'CORN - CHICAGO BOARD OF TRADE',
    'Weizen SRW': 'WHEAT-SRW - CHICAGO BOARD OF TRADE',
    'Weizen HRW': 'WHEAT-HRW - CHICAGO BOARD OF TRADE',
    'Reis': 'ROUGH RICE - CHICAGO BOARD OF TRADE',
    'Hafer': 'OATS - CHICAGO BOARD OF TRADE',
    'Sojabohnen': 'SOYBEANS - CHICAGO BOARD OF TRADE',
    'Sojabohnen Mehl': 'SOYBEAN MEAL - CHICAGO BOARD OF TRADE',
    'Sojabohnen Öl': 'SOYBEAN OIL - CHICAGO BOARD OF TRADE',
    'Raps': 'CANOLA - ICE FUTURES U.S.',
    'Kakao': 'COCOA - ICE FUTURES U.S.',
    'Baumwolle': 'COTTON NO. 2 - ICE FUTURES U.S.',
    'Orangensaft': 'FRZN CONCENTRATED ORANGE JUICE - ICE FUTURES U.S.',
    'Kaffee': 'COFFEE C - ICE FUTURES U.S.',
    'Zucker': 'SUGAR NO. 11 - ICE FUTURES U.S.',
    'Bauholz': 'RANDOM LENGTH LUMBER - CHICAGO MERCANTILE EXCHANGE',
    'Lebendrind': 'LIVE CATTLE - CHICAGO MERCANTILE EXCHANGE',
    'Mastrind': 'FEEDER CATTLE - CHICAGO MERCANTILE EXCHANGE',
    'Magerschwein': 'LEAN HOGS - CHICAGO MERCANTILE EXCHANGE',
    'Butter': 'BUTTER (CASH SETTLED) - CHICAGO MERCANTILE EXCHANGE',
    'Fettarme Milch': 'NON FAT DRY MILK - CHICAGO MERCANTILE EXCHANGE',
    'Milch III': 'MILK, Class III - CHICAGO MERCANTILE EXCHANGE',
    'Milch IV': 'CME MILK IV - CHICAGO MERCANTILE EXCHANGE',
    'Käse': 'CHEESE (CASH-SETTLED) - CHICAGO MERCANTILE EXCHANGE',
    "":""
}

def main(url):
    out = []
    r = requests.get(url).text
    date = url.split('cot-')[1].rstrip('.html')
    l1 = r.split('<tr class="d">')
    for i in l1:
        if '<td>' in i:
            oe1 = '&Ouml;'
            oe2 = '&ouml;'
            ae = '&auml;'
            nt = '&amp;'
            a  = [i for i in i.split('<td>')]
            nm = a[1].strip()
            nm = re.sub('<[^<]+?>', '', nm)

            if nt in nm:
                nm = nm.replace(nt, '&')
            if ae in nm:
                nm = nm.replace(ae, 'ä')
            if oe1 in nm:
                nm = nm.replace(oe1, 'Ö')
            if oe2 in nm:
                nm = nm.replace(oe2, 'ö')

            oi = a[2].strip().rstrip('</td>')
            cl = a[4].strip().rstrip('</td>')
            cs = a[6].strip().rstrip('</td>').lstrip('-')

            if 'Name' in nm:
                pass

            else:
                out.append(dict({
                    "name":names[nm],
                    "date":date,
                    "code":None,
                    "open_interest_all":int(oi),
                    "non_commercial_long":int(cl),
                    "non_commercial_short":int(cs)
                }))

    return out


if __name__ == '__main__':
    data = main('https://cnt1.suricate-trading.de/cotde/history/cot-2022-01-18.html')
    i_fucked_excel(data)