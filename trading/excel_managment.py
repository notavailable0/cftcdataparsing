import pylightxl as xl
import json

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

names2 = {
    "S&P 500 Consolidated - CHICAGO MERCANTILE EXCHANGE": "S&P 500",
    "E-MINI S&P 500 STOCK INDEX - CHICAGO MERCANTILE EXCHANGE": "S&P 500 Mini",
    "DJIA Consolidated - CHICAGO BOARD OF TRADE": "Dow Jones",
    "DOW JONES INDUSTRIAL AVG- x $5 - CHICAGO BOARD OF TRADE": "Dow Jones Mini",
    "NASDAQ-100 Consolidated - CHICAGO MERCANTILE EXCHANGE": "NASDAQ",
    "NASDAQ-100 STOCK INDEX (MINI) - CHICAGO MERCANTILE EXCHANGE": "NASDAQ Mini",
    "E-MINI RUSSELL 2000 INDEX - CHICAGO MERCANTILE EXCHANGE": "RUSSEL 2000",
    "NIKKEI STOCK AVERAGE - CHICAGO MERCANTILE EXCHANGE": "NIKKEI",
    "VIX FUTURES - CBOE FUTURES EXCHANGE": "VIX",
    "U.S. TREASURY BONDS - CHICAGO BOARD OF TRADE": "30 Year Bond",
    "2-YEAR U.S. TREASURY NOTES - CHICAGO BOARD OF TRADE": "2 Year Note",
    "5-YEAR U.S. TREASURY NOTES - CHICAGO BOARD OF TRADE": "5 Year Note",
    "10-YEAR U.S. TREASURY NOTES - CHICAGO BOARD OF TRADE": "10 Year Note",
    "30-DAY FEDERAL FUNDS - CHICAGO BOARD OF TRADE": "Federal Funds",
    "ULTRA U.S. TREASURY BONDS - CHICAGO BOARD OF TRADE": "Ultra U.S. T-Bonds",
    "ULTRA 10-YEAR U.S. T-NOTES - CHICAGO BOARD OF TRADE": "Ultra 10-year U.S. T-Notes",
    "BITCOIN - CHICAGO MERCANTILE EXCHANGE": "BitCoin",
    "EURO FX - CHICAGO MERCANTILE EXCHANGE": "Euro",
    "AUSTRALIAN DOLLAR - CHICAGO MERCANTILE EXCHANGE": "Australischer Dollar",
    "BRAZILIAN REAL - CHICAGO MERCANTILE EXCHANGE": "Brasilianischer Real",
    "BRITISH POUND STERLING - CHICAGO MERCANTILE EXCHANGE": "Britisches Pfund",
    "JAPANESE YEN - CHICAGO MERCANTILE EXCHANGE": "Japanischer Yen",
    "CANADIAN DOLLAR - CHICAGO MERCANTILE EXCHANGE": "Kanadischer Dollar",
    "MEXICAN PESO - CHICAGO MERCANTILE EXCHANGE": "Mexikanischer Peso",
    "NEW ZEALAND DOLLAR - CHICAGO MERCANTILE EXCHANGE": "Neuseeland Dollar",
    "RUSSIAN RUBLE - CHICAGO MERCANTILE EXCHANGE": "Russischer Rubel",
    "SWISS FRANC - CHICAGO MERCANTILE EXCHANGE": "Schweizer Franken",
    "3-MONTH EURODOLLARS - CHICAGO MERCANTILE EXCHANGE": "3 Monate Euro Dollar",
    "U.S. DOLLAR INDEX - ICE FUTURES U.S.": "U.S. Dollar Index",
    "GOLD - COMMODITY EXCHANGE INC.": "Gold",
    "SILVER - COMMODITY EXCHANGE INC.": "Silber",
    "PLATINUM - NEW YORK MERCANTILE EXCHANGE": "Platin",
    "PALLADIUM - NEW YORK MERCANTILE EXCHANGE": "Palladium",
    "COPPER-GRADE #1 - COMMODITY EXCHANGE INC.": "Kupfer",
    "ALUMINUM MW US TR PLATTS - COMMODITY EXCHANGE INC.": "Aluminium",
    "": "Ethanol",
    "COAL (API 2) CIF ARA - NEW YORK MERCANTILE EXCHANGE": "Kohle",
    "CRUDE OIL, LIGHT SWEET - NEW YORK MERCANTILE EXCHANGE": "Rohöl WTI",
    "GASOLINE BLENDSTOCK (RBOB) - NEW YORK MERCANTILE EXCHANGE": "Benzin",
    "#2 HEATING OIL- NY HARBOR-ULSD - NEW YORK MERCANTILE EXCHANGE": "Heizöl",
    "NATURAL GAS - NEW YORK MERCANTILE EXCHANGE": "Erdgas",
    "CHICAGO ETHANOL - NEW YORK MERCANTILE EXCHANGE": "Chicago Ethanol",
    "CORN - CHICAGO BOARD OF TRADE": "Mais",
    "WHEAT-SRW - CHICAGO BOARD OF TRADE": "Weizen SRW",
    "WHEAT-HRW - CHICAGO BOARD OF TRADE": "Weizen HRW",
    "ROUGH RICE - CHICAGO BOARD OF TRADE": "Reis",
    "OATS - CHICAGO BOARD OF TRADE": "Hafer",
    "SOYBEANS - CHICAGO BOARD OF TRADE": "Sojabohnen",
    "SOYBEAN MEAL - CHICAGO BOARD OF TRADE": "Sojabohnen Mehl",
    "SOYBEAN OIL - CHICAGO BOARD OF TRADE": "Sojabohnen Öl",
    "CANOLA - ICE FUTURES U.S.": "Raps",
    "COCOA - ICE FUTURES U.S.": "Kakao",
    "COTTON NO. 2 - ICE FUTURES U.S.": "Baumwolle",
    "FRZN CONCENTRATED ORANGE JUICE - ICE FUTURES U.S.": "Orangensaft",
    "COFFEE C - ICE FUTURES U.S.": "Kaffee",
    "SUGAR NO. 11 - ICE FUTURES U.S.": "Zucker",
    "RANDOM LENGTH LUMBER - CHICAGO MERCANTILE EXCHANGE": "Bauholz",
    "LIVE CATTLE - CHICAGO MERCANTILE EXCHANGE": "Lebendrind",
    "FEEDER CATTLE - CHICAGO MERCANTILE EXCHANGE": "Mastrind",
    "LEAN HOGS - CHICAGO MERCANTILE EXCHANGE": "Magerschwein",
    "BUTTER (CASH SETTLED) - CHICAGO MERCANTILE EXCHANGE": "Butter",
    "NON FAT DRY MILK - CHICAGO MERCANTILE EXCHANGE": "Fettarme Milch",
    "MILK, Class III - CHICAGO MERCANTILE EXCHANGE": "Milch III",
    "CME MILK IV - CHICAGO MERCANTILE EXCHANGE": "Milch IV",
    "CHEESE (CASH-SETTLED) - CHICAGO MERCANTILE EXCHANGE": "Käse",
    "":""
}




def i_fucked_excel(data):
    for i in data:
        if i['name'] in names2.keys(): print('', end='')
        else: print(i['name'])
    dates = []
    daterow = 2
    row_id = 1
    db = xl.Database()
    db.add_ws(ws="Sheet1")
    for i in range(3):
        for dataf in names.keys():
            db.ws(ws="Sheet1").update_index(row=1, col=row_id+1, val=dataf)
            row_id += 1

    #print(data)

    for i in data:
        #print(i)
        dates.append(i['date'])

    dates = list(dict.fromkeys(dates))

    nrow = db.ws(ws='Sheet1').row(row=1)


    for i in dates:
        db.ws(ws="Sheet1").update_index(row=daterow, col=1, val=i)


        for a in data:
            if a['date'] == i:
                ind_clmn = [i for i, n in enumerate(nrow) if n == names2[a['name']]][0]
                db.ws(ws="Sheet1").update_index(row=daterow, col=ind_clmn+1, val=int(str(a['open_interest_all']).replace(",", '')))

        for a in data:
            if a['date'] == i:
                ind_clmn = [i for i, n in enumerate(nrow) if n == names2[a['name']]][1]
                db.ws(ws="Sheet1").update_index(row=daterow, col=ind_clmn+1, val=int(str(a['non_commercial_long']).replace(",", '')))

        for a in data:
            if a['date'] == i:
                ind_clmn = [i for i, n in enumerate(nrow) if n == names2[a['name']]][2]
                print(a['non_commercial_short'])
                print(a['name'])
                try: vl = int(str(a['non_commercial_short']).replace(",", ''))
                except Exception as e: print(e); vl = ''
                db.ws(ws="Sheet1").update_index(row=daterow, col=ind_clmn+1, val=vl)



        daterow += 1

    xl.writexl(db=db, fn="output.xlsx")