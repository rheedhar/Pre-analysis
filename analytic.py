from openpyxl import Workbook, load_workbook
import requests
import openai
import os
from dotenv import load_dotenv

# # create sas file
# data = ["data a; \n", "set a; \n", "label \n"]
# with open("labels.sas", "w") as label_file:
#     label_file.writelines(data)
#
# # API
# load_dotenv()
# openai_endpoint = "https://api.openai.com/v1/completions"
# openai.api_key = os.getenv("OPEN_API_KEY")


wb = load_workbook("test.xlsx")
ws = wb["Labels"]

# TODO: We need to figure out how to detect that a cell is merged. Then we need to unmerge it and distribute the
#  text in the left cell to their respective rows. We do this to the entire document first and then save the document.


merged_cells = ws.merged_cells.ranges
merged_cells_list_string = [str(item) for item in merged_cells]

for merge_cell in merged_cells_list_string:
    print(merge_cell)
    cell_value = ws[merge_cell.split(":")[0]].value
    print(cell_value)
    ws.unmerge_cells(merge_cell)
wb.save("test1.xlsx")

# TODO: Next we need to loop through each row in the worksheet, get the variable name and label, pass label to
#  chatgpt api and then save our result to our sas file.
#
#  for row in ws.iter_rows(min_row=2, min_col=2, max_col=3, values_only=True):
#     exclude_list = ["Patient_ID", "Abstraction_Date"]
#     if any([text in row for text in exclude_list]):
#         new_label = label
#     else:
#         new_label = label.split(" ", 1)[1]
#
#     response = openai.Completion.create(
#         model="text-davinci-003",
#         prompt=f"Create a descriptive label for the following text: '${new_label}'",
#         temperature=0
#     )
#     ai_label = (response["choices"][0]["text"]).strip()
#     with open("labels.sas", "a") as label_file:
#         label_file.write(f"{title}='{ai_label}' \n")
#
# with open("labels.sas", "a") as label_file:
#     label_file.write("run;")
#
#
#
