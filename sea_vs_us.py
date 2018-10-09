import csv
import webbrowser
import gspread
from oauth2client.service_account import ServiceAccountCredentials


'''

    FETCHING DATA FROM GOOGLE SHEETS

'''

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
gc = gspread.authorize(creds)

sheet = gc.open('Kopi2 af Master Data').get_worksheet(0)
china_usa_data = sheet.get_all_records()

sub_categories_us = dict()
sub_categories_china = dict()
sub_categories_sea = dict()

'''

CREATING DICTIONARIES with COMBINED FUNDING ROUNDS FROM CATEGORIES

'''


for line in china_usa_data:
    amount_funded = str(line['USD Conversion'])
    amount_funded = ''.join([i for i in amount_funded if i != ','])
    if amount_funded == '':
        continue
    if not amount_funded[0].isdigit():
        amount_funded = amount_funded[1:]

    try:

        amount_funded = float(amount_funded)
    except ValueError:
        amount_funded = 0

    if line['Country'] == 'China':
        if line['Category'] + ' - ' + line['Sub-category'] not in sub_categories_china:
            sub_categories_china[line['Category'] + ' - ' + line['Sub-category']] = amount_funded
        else:
            sub_categories_china[line['Category'] + ' - ' + line['Sub-category']] += amount_funded

    elif line['Country'] == 'USA':
        if line['Category'] + ' - ' + line['Sub-category'] not in sub_categories_us:
            sub_categories_us[line['Category'] + ' - ' + line['Sub-category']] = amount_funded
        else:
            sub_categories_us[line['Category'] + ' - ' + line['Sub-category']] += amount_funded


sheet = gc.open('Kopi2 af Master Data').get_worksheet(1)

sea_data = sheet.get_all_records()


for line in sea_data:

    amount_funded = str(line['USD Conversion'])

    if amount_funded == '':
        continue

    amount_funded = ''.join([i for i in amount_funded if i != ','])
    if not amount_funded[0].isdigit():
        amount_funded = amount_funded[1:]

    try:
        amount_funded = float(amount_funded)

    except ValueError:
        continue

    if line['Category'] + ' - ' + line['Sub-category'] not in sub_categories_sea:
        sub_categories_sea[line['Category'] + ' - ' + line['Sub-category']] = amount_funded
    else:
        sub_categories_sea[line['Category'] + ' - ' + line['Sub-category']] += amount_funded



'''

    INVESTMENT SCORE CALCULATION (needs cleaning)

'''


us_total = sum(sub_categories_us.values())
china_total = sum(sub_categories_china.values())
sea_total = sum(sub_categories_sea.values())
scores_china = dict()
scores_us = dict()



for i in sub_categories_china:
    if i in sub_categories_sea:
        k = sub_categories_china[i]/china_total
        ratio = ((sub_categories_china[i]/china_total)/(sub_categories_sea[i]/sea_total))
        #except ZeroDivisionError:
        #    ratio = 0
        scores_china[i] = ratio
        '''

        if k < 0.001:
            scores_china[i] = 0.0000000000001
        else:
            scores_china[i] = ratio


        SEA = (sub_categories_sea[i]/sea_total)

        if SEA < 0.001:
            scores_china[i] = 0.0000000000001
        else:
            scores_china[i] = ratio
#    else:
#        scores_china[i] = 0

        '''

for i in sub_categories_us:
    if i in sub_categories_sea:
        k = sub_categories_us[i]/us_total

        #try:
        ratio = ((sub_categories_us[i]/us_total)/(sub_categories_sea[i]/sea_total))
        #except ZeroDivisionError:
        #    ratio = 0
        scores_us[i] = ratio
        '''
        if k < 0.00005 or SEA < 0.001:
            scores_us[i] = 0.0000000000001

        else:
            scores_us[i] = ratio

        SEA = (sub_categories_sea[i]/sea_total)

        if SEA < 0.001:
            scores_us[i] = 0.0000000000001
        else:
            scores_us[i] = ratio
#    else:
#        scores_us[i] = 0
        '''

#for i in scores_china.keys():
#    if isinstance(str, scores_china[i]):
    #    print(scores_china[i


china_top = sorted(scores_china, key=scores_china.get, reverse=True)
us_top = sorted(scores_us, key=scores_us.get, reverse=True)

writing_list = []


for cat in china_top:
    cats = cat.split(' - ')
    try:
        writing_list.extend((None, cats[0], cats[1], '', '', scores_china[cat], sub_categories_china[cat]/china_total*100, sub_categories_sea[cat]/sea_total*100))
    except KeyError:
        writing_list.extend((None, cats[0], cats[1], '', '', scores_china[cat], sub_categories_china[cat]/china_total*100, 'Category not in SEA!'))



writing_list_us = []

for cat in us_top:
    cats = cat.split(' - ')
    try:
        writing_list_us.extend((None, cats[0],cats[1], '', '', scores_us[cat], sub_categories_us[cat]/us_total*100, sub_categories_sea[cat]/sea_total*100))
    except KeyError:
        writing_list_us.extend((None, cats[0],cats[1], '', '', scores_us[cat], sub_categories_us[cat]/us_total*100, 'Category not in SEA!'))

        pass


'''

    ADDING TO NEW GOOGLE SHEETS SPREADSHEET

'''


sheet_china = gc.open('Kopi2 af Master Data').get_worksheet(4) # SEA data
sheet_us = gc.open('Kopi2 af Master Data').get_worksheet(5) #US China India Data


'''
sheet_china = sh.add_worksheet(title = 'Fund. Diff. China to SEA', rows =len(writing_list)+10 , cols = 10)
sheet_usa = sh.add_worksheet(title = 'Fund. Diff. USA to SEA', rows =len(writing_list)+10 , cols = 10)
'''


cell_list = sheet_china.range('A2:H' + str(int(len(writing_list)/8 + 1)))
cell_list_us = sheet_us.range('A2:H' + str(int(len(writing_list_us)/8 +1)))



for i, cell in enumerate(cell_list):
    cell.value = writing_list[i]

for i, cell in enumerate(cell_list_us):
    cell.value = writing_list_us[i]


sheet_china.update_cells(cell_list)
sheet_us.update_cells(cell_list_us)
