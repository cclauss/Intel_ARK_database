# coding: utf-8

import datetime, json

filename = 'Intel_ARK_AdvancedSearch_2015_08_31.json'

print("json.load('{}')".format(filename))
with open(filename) as in_file:
    start = datetime.datetime.now()
    intel_dict = json.load(in_file)
    print('Elapsed time: {}'.format(datetime.datetime.now() - start))  # 5 sec

# [u'@xmlns', u'@xmlns:ss', u'ss:Styles', u'ss:Worksheet']
print(sorted(intel_dict['ss:Workbook'].keys()))
print(intel_dict['ss:Workbook']['ss:Worksheet']['@ss:Name'])  # Sheet1
# ARK | Compare Intel® Products
# 8/31/2015 11:35:35 AM
# [processor list]
rows = intel_dict['ss:Workbook']['ss:Worksheet']['ss:Table']['ss:Row']
for x in rows[:3]:
    cell = x['ss:Cell']
    if isinstance(cell, dict):
        data = cell['ss:Data']
        print(data.get('text', data.get('#text', None)))
    else:
        for i, item in enumerate(cell):
            if item:
                print('{:>4} {}'.format(i, item['ss:Data']['#text']))
    print('')
