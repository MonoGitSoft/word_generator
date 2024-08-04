
import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl

parser = argparse.ArgumentParser(description='Generate words documents based on a templated word doc and data from excel file.')

parser = argparse.ArgumentParser(add_help=False)

parser.add_argument('-h', '--help', action='help', default=argparse.SUPPRESS,
                    help='Show this help message and exit.')

parser.add_argument('-w', '--word', metavar='', help='Name of the word document that contains the placeholders.', default="meghatalmazas.docx")
parser.add_argument('-o', '--output', metavar='', help='Where to generate the word docs', default='OUTPUT')

args = parser.parse_args()
print(args.output)

base_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
word_template_path = base_dir / args.word
excel_path = base_dir / "contracts-list.xlsx"
output_dir = base_dir / "OUTPUT"

# Create output folder for the word documents
output_dir.mkdir(exist_ok=True)

# Convert Excel sheet to pandas dataframe
df = pd.read_excel(excel_path, sheet_name="Sheet1")

df = df.fillna('')

df["TODAY"] = datetime.today().strftime('%Y-%m-%d')

# Iterate over each row in df and render word document
for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template_path)
    doc.render(record)
    output_path = output_dir / f"{record['FIRST_NAME']}-{record['LAST_NAME']}-{args.word}.docx"
    doc.save(output_path)
