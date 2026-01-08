import pandas as pd
from lxml import etree as ET
from datetime import datetime
import sys
import os

# ----------------------------
# Get Excel file from argument
# ----------------------------
if len(sys.argv) < 2:
    raise Exception("Excel file name not provided")

excel_file = sys.argv[1]

if not os.path.exists(excel_file):
    raise Exception(f"File not found: {excel_file}")

# ----------------------------
# Load Excel file
# ----------------------------
df = pd.read_excel(excel_file, sheet_name=0)

# Drop unnamed columns
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

# ----------------------------
# Create XML structure
# ----------------------------
root = ET.Element("testcases")

for index, row in df.iterrows():
    testcase = ET.SubElement(root, "testcase", {
        "name": str(row.get("Testcase ID", index + 1))
    })

    custom_fields = ET.SubElement(testcase, "custom_fields")

    for col in df.columns:
        name = str(col)
        value = str(row[col]) if pd.notna(row[col]) else ""

        cf = ET.SubElement(custom_fields, "custom_field")

        name_elem = ET.SubElement(cf, "name")
        name_elem.text = ET.CDATA(name)

        value_elem = ET.SubElement(cf, "value")
        value_elem.text = ET.CDATA(value)

# ----------------------------
# Output XML file
# ----------------------------
base_name = os.path.splitext(os.path.basename(excel_file))[0]
today_date = datetime.now().strftime("%Y-%m-%d")
output_file = f"{base_name}_{today_date}.xml"

tree = ET.ElementTree(root)
tree.write(
    output_file,
    encoding="utf-8",
    xml_declaration=True,
    pretty_print=True
)

print(f"Success! XML generated: {output_file}")
