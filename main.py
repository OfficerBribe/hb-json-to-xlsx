"""
Outputs xlsx file from HumbleBundle's store page's json database
You can get the database by looking at their website's source code

xlsxwriter - Copyright 2013-2015, John McNamara, jmcnamara@cpan.org
"""

# encoding: utf-8
import xlsxwriter
import json

# Headers in spreadsheet
heads_h = ['Title',
           'Price EUR',
           'Full Price EUR',
           '% off',
           # Platforms
           'Windows',
           'Linux',
           'Mac',
           'Android',
           'MP3',
           'FLAC',
           # Delivery method
           'Steam',
           'DRM-free',
           'GOG',
           'Uplay',
           'Other key',
           'Desura',
           'asmjs']

# Appropriate json tags based from heads_h
heads_m = ['human_name',
           'current_price',
           'full_price',
           # not yet created
           'percent_off']

# nested inside 'platforms'
platforms = ['windows',
             'linux',
             'mac',
             'android',
             'mp3',
             'flac']

# nested inside 'delivery_methods'
delivery_methods = ['steam',
                    'download',
                    'gog',
                    'uplay'
                    'other-key',
                    'desura',
                    'asmjs']

# Parent tags for [win, linux, mac] and [steam, downloads]
heads_m2 = ['platforms',
            'delivery_methods']

# Keep track of how many win, linux, mac, steam, etc. in data
counts = {}

# For more readable code. item[price] VS item[heads_m[1]]
title = heads_m[0]
price = heads_m[1]
full_price = heads_m[2]
percent_off = heads_m[3]

# Open json
data = json.load(open('data.json'))

# Test what platforms or delivery methods are in data
#methods = set()
# for item in data:
#    if heads_m2[1] in item:
#        for method in item[heads_m2[1]]:
#            methods.add(method)
# print(methods)

# Create xlsx and a new worksheet
workbook = xlsxwriter.Workbook('HB_sale.xlsx')
worksheet = workbook.add_worksheet()

# Determine max column size for Title
col_a_size = max(len(item[title]) for item in data if price in item)

# Format column sizes, add text formating variables
bold = workbook.add_format({'bold': True})
worksheet.set_column('A:A', col_a_size)
worksheet.set_column('B:B', 10)
worksheet.set_column('C:C', 12)
worksheet.set_column('E:E', 12)
worksheet.set_column('L:L', 12)
worksheet.set_column('N:N', 10)

# Write headers
for head in enumerate(heads_h):
    worksheet.write(0, head[0], head[1], bold)

# Calculate %off and add to data
for item in data:
    # Only games
    if 'game' in item.get('content_types', []):
        item[percent_off] = round((item[full_price][0] -
                                   item[price][0]) * 100 /
                                  item[full_price][0])
        item[full_price] = item[full_price][0]
        item[price] = item[price][0]

# Write from data to spreadsheet
row = 1
for item in data:
    col = 0
    if 'game' in item.get('content_types', []):
        # Write rows with text/numbers
        for head in heads_m:
            worksheet.write(row, col, item[head])
            col += 1

        # Write rows with bool yes/no for platforms
        for head in platforms:
            if head in item[heads_m2[0]]:
                worksheet.write(row, col, 'X')
                counts[head] = counts.get(head, 0) + 1
            col += 1

        # Write rows with bool yes/no for delivery methods
        for head in delivery_methods:
            if head in item[heads_m2[1]]:
                worksheet.write(row, col, 'X')
                counts[head] = counts.get(head, 0) + 1
            col += 1
        row += 1
        counts['total'] = counts.get('total', 0) + 1

workbook.close()

print('--Statistics--')
print('Total games:', str(counts.get('total', 0)))
print('* ', 'Windows:', str(counts.get('windows', 0)))
print('* ', 'Linux:', str(counts.get('linux', 0)))
print('* ', 'Mac:', str(counts.get('mac', 0)))
print('* ', 'Android:', str(counts.get('android', 0)))
print('Delivery methods:')
print('* ', 'Steam:', str(counts.get('steam', 0)))
print('* ', 'DRM-free:', str(counts.get('download', 0)))
print('* ', 'GOG:', str(counts.get('gog', 0)))
print('* ', 'Other key:', str(counts.get('other-key', 0)))
print('* ', 'Desura:', str(counts.get('desura', 0)))
print('* ', 'asmjs:', str(counts.get('asmjs', 0)))
print('Soundtracks:')
print('* ', 'mp3:', str(counts.get('mp3', 0)))
print('* ', 'flac:', str(counts.get('flac', 0)))
