import re
import os
import json
from docxtpl import DocxTemplate, RichText
from core import utils

def main():
    with open('config.json') as f:
        config = json.load(f)

    document = DocxTemplate(config['template'])

    # Initialize variables
    assessment = {}
    components = []

    # Import assessment result from directory
    for file_name in os.listdir(config['assessment']):
        file_path = '{base}/{file}'.format(base=config['assessment'], file=file_name)
        library_path = config['library']

        if (component := utils.find_matching_key(config['components'], file_name)) is not None:
            tmpd = {}
            tmpd['name'] = config['components'][component]
            tmpd['shortname'] = component
            
            print("[*] Creating assessment,", tmpd['shortname'])

            tmpd['findings'] = utils.sort_by_risk(utils.unflatten(file_path, library_path, config['group']), config['risk_matrix'])
            components.append(tmpd)

    assessment['components'] = components
    assessment['report_config'] = config['report_configuration']
    
    # For debugging
    # print(assessment, file=open("debug.json", "w", encoding='utf-8'))

    utils.generate_report(config['template'], assessment, config['report_name'])
    print("[*] Report has been generated:", config['report_name'])

if __name__ == '__main__':
    main()