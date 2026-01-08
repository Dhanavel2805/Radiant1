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
# Output XML file (Save in 2 places)
# ----------------------------
base_name = os.path.splitext(os.path.basename(excel_file))[0]
today_date = datetime.now().strftime("%Y-%m-%d")
output_file_name = f"{base_name}_{today_date}.xml"

# Get Excel directory
excel_dir = os.path.dirname(os.path.abspath(excel_file))

# Path 1: Same directory
output_path_1 = os.path.join(excel_dir, output_file_name)

# Path 2: Result folder
result_dir = os.path.join(excel_dir, "Result")
os.makedirs(result_dir, exist_ok=True)
output_path_2 = os.path.join(result_dir, output_file_name)

tree = ET.ElementTree(root)

# Write to both locations
tree.write(
    output_path_1,
    encoding="utf-8",
    xml_declaration=True,
    pretty_print=True
)

tree.write(
    output_path_2,
    encoding="utf-8",
    xml_declaration=True,
    pretty_print=True
)

print("Success! XML generated in two locations:")
print(f"1️⃣ {output_path_1}")
print(f"2️⃣ {output_path_2}")
