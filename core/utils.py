import re
import math
import pandas as pd
from operator import itemgetter
from docxtpl import DocxTemplate

# Import issue library
def read_excel(file, sheet_name=0, header=0, type="record"):
    df = pd.read_excel('{file}'.format(file=file), sheet_name=sheet_name, header=header, engine='openpyxl')
    if type == "list":
        return df.set_index("Issue Title").T.to_dict()
    else:
        return df.to_dict(orient='records')

def find_matching_key(lists, keyword):
    for comp in lists:
        if comp in keyword:
            return comp
    return None

def generate_report(template, content, output):
    document = DocxTemplate(template)
    document.render(content)
    document.save(output)

def unflatten(file, library, configuration):
    tmp = {}
    unique = []
    
    content = read_excel(file, sheet_name=0, header=0)
    definition = read_excel(library, header=0, type="list")

    # Create list of dict with array
    for item in content:
        tmp[item['Issue Title']] = {}
        tmp[item['Issue Title']]['affected'] = []

    count = 0
    for item in content:
        tmpk = {}
  
        for k, v in item.items():
            if k in configuration:
                tmpk[clean_string(k)] = v
            if k not in configuration:
                if k == "Issue Title":
                    if isinstance(definition[item['Issue Title']]['New Title'], float):
                        tmp[item['Issue Title']][clean_string(k)] = item['Issue Title']
                    else:
                        tmp[item['Issue Title']][clean_string(k)] = definition[item['Issue Title']]['New Title']
                else:
                    if isinstance(definition[item['Issue Title']][k], float):
                        tmp[item['Issue Title']][clean_string(k)] = ""
                    else:
                        tmp[item['Issue Title']][clean_string(k)] = definition[item['Issue Title']][k]

        if "\"" in item['Affected File(s)']:
            sheet_name = re.findall(r'"([^"]*)"', item['Affected File(s)'])
            tmp[item['Issue Title']]['affected'] = retrieve_list(file, sheet_name, configuration)
        else:
            tmp[item['Issue Title']]['affected'].append(tmpk)
        
        count = count + len(tmp[item['Issue Title']]['affected'])
    
    for k, v in tmp.items():
        unique.append(v)

    # For debugging
    # for i in unique:
    #     i = {k: str(v).encode("utf-8") for k,v in i.items()}
    #     print("{comp}:{x}".format(comp=i['issuetitle'],x=i['affected']))

    print("[+] Added {count} records".format(count=count))

    return unique

def retrieve_list(file, sheet_name, configuration):
    list_of_items = []
    sheet = read_excel(file, sheet_name=sheet_name[0], header=1)

    for item in sheet:
        tmp = {}
        for k, v in item.items():
            if k in configuration:
                tmp[clean_string(k)] = v
        list_of_items.append(tmp)
    return list_of_items


def retrieve_definition(file, issue):
    return issue

def print_list(content, key):
    for item in content:
        print(item[key])

def clean_string(word):
    setting = str.maketrans({'(':None, ')':None, ' ':None})
    return word.lower().translate(setting)

def sort_by_risk(content, definition):
    # Sort definition by priority
    order = sorted(definition, key=lambda x: definition[x]['priority'], reverse=True)
    risk = {key: definition.get(key) for key in order}
    
    tmp = content
    for item in risk:
        tmp = sorted(tmp, key=lambda x: risk[item]['order'].index(x[item]))

    for item in tmp:
        if item['risk'] == "Critical": item['background'] = '5f497a'
        if item['risk'] == "High": item['background'] = 'c00000'
        if item['risk'] == "Medium": item['background'] = 'e36c0a'
        if item['risk'] == "Low": item['background'] = '339900'
        if item['risk'] == "Informational": item['background'] = '0070c0'
    return tmp