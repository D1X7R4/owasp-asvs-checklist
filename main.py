from json import loads
from worksheet_generator import WorkSheetGenerator
from overview_worksheet import OverviewWorksheet
from xlsxwriter import Workbook
import argparse

parser =  argparse.ArgumentParser()
parser.add_argument('-i', '--input-json', help='JSON input file with the ASVS controls', required=True)
parser.add_argument('-o', '--output', help='Output filename')
args = parser.parse_args()

def readCSV(filename):
    f = open(filename, "r")
    data = loads(f.read())
    f.close()
    return data

if __name__ == "__main__":
    data = readCSV(args.input_json)
    if not args.output:
        args.output = f'ASVS-Checklist-v{data["Version"]}'

    workbook = Workbook(f'{args.output}.xlsx')

    OverviewWorksheet(workbook=workbook,worksheet_title="Overview", categories=[(x['Name'], x['ShortName']) for x in data['Requirements']])

    for category in data['Requirements']:
        worksheet = WorkSheetGenerator(workbook=workbook,worksheet_title=data['Name'], worksheet_name=category["Name"], worksheet_shortname=category["ShortName"], name=data["ShortName"], version=data["Version"])
        worksheet.generateWorksheet(category['Items'])

    workbook.close()