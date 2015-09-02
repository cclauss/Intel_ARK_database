# coding: utf-8

import datetime, json, xmltodict

# xmltodict.parse(file.xml.xls) takes 2 min 30 sec on my 32-bit iPad
# json.dump(intel_dict, file.json) takes 30 sec
# json.load(file.json) takes 5 sec -- See Intel_json_reader.py

filename = 'Intel_ARK_AdvancedSearch_2015_08_31.xml.xls'
out_filename = filename.partition('.')[0] + '.json'

print("xmltodict.parse('{}')".format(filename))
with open(filename) as in_file:
    in_file.read(2)  # remove the first two bytes of junk
    start = datetime.datetime.now()
    intel_dict = xmltodict.parse(in_file)
    print('Elapsed time: {}'.format(datetime.datetime.now() - start))  # 2 min 30 sec

if intel_dict:
    with open(out_filename, 'w') as out_file:
        print("\njson.dump('{}')".format(out_filename))
        start = datetime.datetime.now()
        json.dump(intel_dict, out_file)
        print('Elapsed time: {}'.format(datetime.datetime.now() - start)) # 30 sec
else:
    exit('intel_dict is empty!')
print('Done.')
