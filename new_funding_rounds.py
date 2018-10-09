import urllib.request
from datetime import datetime
from xlrd import open_workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials

#Crunchbase Url download to List of dictionaries.

user_key = '1b4df7023be98d6b272fd9651e63de31'

url = 'https://api.crunchbase.com/v3.1/excel_export/crunchbase_export.xlsx?user_key=' + user_key

urllib.request.urlretrieve(url, "crunchbase_export.xlsx")


excel_list = []
book = open_workbook('crunchbase_export.xlsx')
sheet = book.sheet_by_index(2)

# read first row for keys
keys = sheet.row_values(0)
# read the rest rows for values
values = [sheet.row_values(i) for i in range(1, sheet.nrows)]

for value in values:
    excel_list.append(dict(zip(keys, value)))

print('Converted Excel to format')

'''
    Google initialization and find Worksheets (Master Data)
'''
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
gc = gspread.authorize(creds)

sheet = gc.open('Kopi2 af Master Data').get_worksheet(1) # SEA data
sheet_us_china = gc.open('Kopi2 af Master Data').get_worksheet(0) #US China India Data



sea_data = sheet.get_all_records()
us_china_data = sheet_us_china.get_all_records()

print('Google initialization Done')

'''
Indi. company data
'''

company_sheet = book.sheet_by_index(1)
keys = company_sheet.row_values(0)
values = [company_sheet.row_values(i) for i in range(1, company_sheet.nrows)]

comp_list = []
for value in values:
    comp_list.append(dict(zip(keys,value)))

company_desc_dict = dict()
company_domains = dict()
for company in comp_list:
    company_desc_dict[company['company_name']] = company['short_description']
    company_domains[company['company_name']] = company['domain']
    #founders[company['company_name']] = company['']



count = 0


date_today = excel_list[0]['announced_on']
SEA_country_codes = {'THA': 'Thailand', 'IDN': 'Indonesia', 'VNM': 'Vietnam','PHL' : 'Phillipines', 'MYS':'Malaysia','SGP':'Singapore', 'MMR':'Myanmar', 'KHM' : 'Cambodia'}
oth_country_codes = {'CHN': 'China', 'IND': 'India', 'USA': 'USA'}

# Iterates through new funding rounds given todays date




def get_desc(company):
    return company_desc_dict[company]

def get_site(domain):
    if 'http://' + company_domains[company] == 'http://':
        return ''
    else:
        return 'http://' + company_domains[company]

def date_crunch_to_master(old):
    dt = datetime.strptime(old, '%Y-%m-%d')
    return '{0}/{1}/{2:02}'.format(dt.month, dt.day, dt.year)


sea_writing_list = []
us_writing_list = []


def get_inv_lead(investors):
    inv_list = investors.split(', ')
    inv, lead = [], []
    for i in inv_list:
        if 'Lead' == i[0:4]:
            lead.append(i.split(' - ')[-1])
        else:
            inv.append(i)
    return inv, lead



usa_investors = set(['Sequoia Capital','Accel Partners','Benchmark','Andreessen Horowitz','First Round Capital','Felicis Ventures','Shasta Ventures','Google','Founders Fund','Kleiner Perkins Caufield & Byers','Battery Ventures','GGV Capital','NEA','Greylock Partners','Lightspeed Venture Partners','Bessemer Venture Partners','General Catalyst','Social Capital','Spark Capital','Union Square Ventures'])
china_investors = set(['Sequoia Capital','GSR Ventures','Tencent Holdings','Morningside Venture Capital','Bertelsmann Asia Investment Fund','ZhenFund','K2VC','Qiming Venture Partners','IDG Capital Partners','SB China Capital','Softbank China Venture Capital','Legend Capital','Shenzhen Capital Group','Capital Today','Fortune Capital','NewMargin Ventures','SAIF Partners','Baidu','Alibaba','Alibaba Capital Partners','Alibaba Group'])
india_investors = set(['Blume Ventures','Indian Angel Network','Sequoia Capital','Sequoia Capital India','Accel Partners','Mumbai Angels','Kalaari Capital','IDG Ventures India','IDG Ventures','Tiger Global Management','SAIF Partners','Helion Venture Partners','Orios Venture Partners','Nexus Venture Partners','Kae Capital','Goldman Sachs','Bain Capital Ventures'])


print('Beginning Loop')

for fr in excel_list:
    if fr['announced_on'] != date_today:
        break

    #for sea

    if fr['country_code'] in SEA_country_codes:
        company = fr['company_name']
        company_found = False

        for round in sea_data:
            if round['Organization Name'] == company:
                company_found = True
                temp = round
                break

        if company_found:
            cat, subcat, desc, website, founder = temp['Category'], temp['Sub-category'], temp['Description'], temp['Website'], temp['Founder(s)']
        else:
            cat, subcat, desc, website, founder = '', '', get_desc(company), get_site(company), ''

        try:
            usd_conversion = fr['raised_amount_usd']
            MM = float(fr['raised_amount_usd'])/1000000
        except ValueError:
            usd_conversion = 'Undisclosed'
            MM = 'Undisclosed'
        inv, lead = get_inv_lead(fr['investor_names'])

        line = [fr['company_name'], cat, subcat, '', desc, fr['raised_amount_currency_code'], MM, usd_conversion, '',  fr['funding_round_type'], date_crunch_to_master(fr['announced_on']), lead, inv, fr['cb_url'], SEA_country_codes[fr['country_code']], founder, website]
        sea_writing_list.extend(line)
        print('Added Sea line')
        sheet.append_row(line)

    # FOR US INDIA CHINA

    if fr['country_code'] in oth_country_codes:

        company = fr['company_name']
        company_found = False
        inv, lead = get_inv_lead(fr['investor_names'])
        inv.extend(lead)

        # Check if investor of interest
        if fr['country_code'] == 'USA':
            print(company)
            if len(set(inv).intersection(usa_investors)) == 0:
                continue

        if fr['country_code'] == 'CHN':

            if len(set(inv).intersection(china_investors)) == 0:
                continue

        if fr['country_code'] == 'IND':
            if len(set(inv).intersection(india_investors)) == 0:
                continue




        for round in us_china_data:
            if round['Organization Name'] == company:
                company_found = True
                temp = round
                break

        if company_found:
            try:
                cat, subcat, desc, website, founder, founded_date = temp['Category'], temp['Sub-category'], temp['Description'], temp['Website'], temp['Founder(s)'], temp['Founded Date']

            except KeyError:
                cat, subcat, desc, website, founder, founded_date = temp['Category'], temp['Sub-category'], temp['Description'], temp['Website'], '', temp['Founded Date']
        else:
            cat, subcat, desc, website, founder,founded_date = '', '', get_desc(company), get_site(company), '', ''

        try:
            usd_conversion = fr['raised_amount_usd']
            MM = float(fr['raised_amount_usd'])/1000000
            ass = 'No'
        except ValueError:
            usd_conversion = 'Undisclosed'
            MM = 'Undisclosed'
            ass = ''
        inv, lead = get_inv_lead(fr['investor_names'])

        #WRITE LINE IN MISSING AND CHECK IT WORK
        line = [company, cat, subcat, desc, website, oth_country_codes[fr['country_code']], date_crunch_to_master(fr['announced_on']), fr['funding_round_type'], usd_conversion, ass,   founded_date]
        print('added us line')

        sheet_us_china.append_row(line)

        # for China, India, SEA
