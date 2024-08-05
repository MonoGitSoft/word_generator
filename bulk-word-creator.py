
import argparse
import glob
import os
from datetime import datetime
from pathlib import Path
import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl

def get_docx_files(directory):
    pattern = os.path.join(directory, '*.docx')
    docx_files = glob.glob(pattern)
    return docx_files

parser = argparse.ArgumentParser(description='Generate words documents based on a templated word doc and data from excel file.')

parser = argparse.ArgumentParser(add_help=False)

parser.add_argument('-h', '--help', action='help', default=argparse.SUPPRESS,
                    help='Show this help message and exit.')

parser.add_argument('-w', '--word', metavar='', help='Name of the word document that contains the placeholders.', default="meghatalmazas.docx")
parser.add_argument('-o', '--output', metavar='', help='Where to generate the word docs', default='OUTPUT')

args = parser.parse_args()
print(args.output)

base_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
word_templates_path = base_dir / "word_templates"
excel_path = base_dir / "contracts-list.xlsx"
output_dir = base_dir / "OUTPUT"

# Create output folder for the word documents
output_dir.mkdir(exist_ok=True)


# Convert Excel sheet to pandas dataframe
df = pd.read_excel(excel_path, sheet_name="Sheet1")

df = df.fillna('')

df["TODAY"] = datetime.today().strftime('%Y-%m-%d')

df["BIRTH_TIME"] = pd.to_datetime(df["BIRTH_TIME"])
df["BIRTH_TIME_Y"] = df["BIRTH_TIME"].dt.year
df["BIRTH_TIME_M"] = df["BIRTH_TIME"].dt.month
df["BIRTH_TIME_D"] = df["BIRTH_TIME"].dt.day
df["BIRTH_TIME"] = pd.to_datetime(df["BIRTH_TIME"]).dt.date



df["PASSPORT_VALID_DATE"] = pd.to_datetime(df["PASSPORT_VALID_DATE"])
df["PASSPORT_VALID_DATE_Y"] = df["PASSPORT_VALID_DATE"].dt.year
df["PASSPORT_VALID_DATE_M"] = df["PASSPORT_VALID_DATE"].dt.month
df["PASSPORT_VALID_DATE_D"] = df["PASSPORT_VALID_DATE"].dt.day
df["PASSPORT_VALID_DATE"] = pd.to_datetime(df["PASSPORT_VALID_DATE"]).dt.date



# Iterate over each row in df and render word document
for record in df.to_dict(orient="records"):
    for template in get_docx_files(word_templates_path):
        print(os.path.basename(template))
        doc = DocxTemplate(template)
        doc.render(record)
        output_dir_person = output_dir / f"{record['FIRST_NAME']}-{record['LAST_NAME']}"
        output_dir_person.mkdir(exist_ok=True)
        output_path = output_dir_person / f"{record['FIRST_NAME']}-{record['LAST_NAME']}-{os.path.basename(template)}"
        doc.save(output_path)

print(df.dtypes)