import xml.etree.ElementTree as ET
import pandas as pd
import os
import re
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# directory containing XML files
dir_path = "PATH XML FILES"

# initialize data frame to store results
results = pd.DataFrame(columns=["VLDTN_ID", "Frequency", "Date", "Date of Validation"])

# loop through XML files in directory
for filename in os.listdir(dir_path):
    if filename.endswith(".xml"):
        file_path = os.path.join(dir_path, filename)

        # parse XML file
        tree = ET.parse(file_path)
        root = tree.getroot()

        # count frequency of every "VLDTN_ID" attribute
        vldtn_id_counts = {}
        for elem in root.iter():
            if "VLDTN_ID" in elem.attrib:
                vldtn_id = elem.attrib["VLDTN_ID"][:6]
                if vldtn_id in vldtn_id_counts:
                    vldtn_id_counts[vldtn_id] += 1
                else:
                    vldtn_id_counts[vldtn_id] = 1

        # extract date from filename
        date_match = re.search(r"_(\d{6})_(\d{8})", filename)
        if date_match:
            date = date_match.group(1)
            date_of_validation = date_match.group(2)
        else:
            date = ""
            date_of_validation = ""

        # add results to data frame
        file_results = pd.DataFrame(list(vldtn_id_counts.items()), columns=["VLDTN_ID", "Frequency"])
        file_results["Date"] = date
        file_results["Date of Validation"] = date_of_validation
        results = pd.concat([results, file_results], ignore_index=True)

# write results to an Excel file
results.to_excel("PATH RESULT EXCEL", index=False, sheet_name="Sheet1", startrow=0, header=True)

# create a new data frame for sum of "VLDTN_ID" frequencies
vldtn_id_sum = results.groupby("VLDTN_ID").sum().reset_index()
vldtn_id_sum = vldtn_id_sum[["VLDTN_ID", "Frequency"]]

# calculate the percentage of the whole frequency for every row in column 3
total_frequency = vldtn_id_sum["Frequency"].sum()
vldtn_id_sum["Percentage"] = (vldtn_id_sum["Frequency"] / total_frequency) * 100

# write sum results to a new sheet in the same Excel file
with pd.ExcelWriter("PATH RESULT EXCEL", mode="a") as writer:
    vldtn_id_sum.to_excel(writer, index=False, sheet_name="Sheet2", startrow=0, header=True)

