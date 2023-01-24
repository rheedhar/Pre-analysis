from openpyxl import Workbook, load_workbook
import requests
import openai
import os
from dotenv import load_dotenv

# create sas file
data = ["data a; \n", "set a; \n", "label \n"]
with open("labels.sas", "w") as label_file:
    label_file.writelines(data)

# API
load_dotenv()
openai_endpoint = "https://api.openai.com/v1/completions"
openai.api_key = os.getenv("OPEN_API_KEY")


wb = load_workbook("test.xlsx")
ws = wb["Labels"]

# loop through rows in worksheet
for row in ws.iter_rows(min_row=2, min_col=2, max_col=3, max_row=6, values_only=True):
    title, label = row
    exclude_list = ["Patient_ID", "Abstraction_Date"]
    if any([text in row for text in exclude_list]):
        new_label = label
    else:
        new_label = label.split(" ", 1)[1]

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=f"Create a descriptive label for the following text: '${new_label}'",
        temperature=0
    )
    ai_label = (response["choices"][0]["text"]).strip()
    with open("labels.sas", "a") as label_file:
        label_file.write(f"{title}='{ai_label}' \n")

with open("labels.sas", "a") as label_file:
    label_file.write("run;")



