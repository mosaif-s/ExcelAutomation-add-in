import pyautogui
import time
import xlwings as xw
from openai import OpenAI
import os
from dotenv import load_dotenv

load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

#Setup directory for flag files
os.makedirs("C:/temp", exist_ok=True)
with open("C:/temp/flag.txt", "w") as f:
    f.write("-")

#Setup the excel sheet into a doubly linked list
wb = xw.apps.active.books.active
ws = wb.sheets.active

top_left = "A1"
last_cell = ws.used_range.last_cell
data = ws.range(f"{top_left}:{last_cell.address}").value

with open("C:/temp/show_loading.txt", "w") as f:
    f.write("showModal")

for i in range(len(data)):
    for j in range(len(data[i])):
        if data[i][j] is None:
            letter=chr(ord('A') + j)
            data[i][j]=f"{letter}{i+1}"
        else:
            temp=data[i][j]
            letter=chr(ord('A') + j)
            data[i][j]=f"{letter}{i+1}={temp}"

print(data)
as_text = '\n'.join(['\t'.join([str(cell) if cell is not None else '' for cell in row]) for row in data])

#Get output from LLM
client = OpenAI(api_key=OPENAI_API_KEY)
response = client.chat.completions.create(
    model="gpt-4o",  # <-- Changed from gpt-3.5-turbo
    messages=[
        {
            "role": "system",
            "content": (
                "You are a command-only agent.\n"
                "You may use the following commands as many times as necessary, in any order:\n"
                "Click <cell>\n"
                "Enter <text>\n"
                "Press ENTER\n\n"
                "Do NOT include explanations, punctuation, markdown, or any formatting.\n"
                "Each command must be on a separate line.\n"
                "No quotation marks around text or cell references.\n"
                "Only output raw commands.\n\n"
                "You will be given a nested list representing an Excel worksheet.\n"
                "Only use Excel-compatible functions (NOT Google Sheets functions).\n"
                "DO NOT use FLATTEN, ARRAYFORMULA, or any other Google Sheets-specific formulas.\n"
                "Your goal is to fill in the cell containing 'FILL THIS' with the appropriate value or formula, based on the context of the surrounding cells.\n"
                "When you see a cell with the text 'FILL THIS...' treat it as an instruction to act on, not as content to replicate. Extract the meaning of the text and follow the instruction accordingly.\n"
                "Assume the sheet is being edited in Excel, not Google Sheets."
            )
        },
        {
            "role": "user",
            "content": f"{data}"
        }
    ],
    temperature=0.2
)
print(response.choices[0].message.content)

ai_output=response.choices[0].message.content
#ai_output="Click B6\nEnter =SUM(B2:B5)\nPress ENTER"
ai_output_list=ai_output.split("\n")

# Show flags
with open("C:/temp/result.txt", "w") as f:
    f.write(str(len(ai_output_list)))

with open("C:/temp/done.txt", "w") as f:
    f.write("done")

while os.path.exists("C:/temp/flag.txt"):
    time.sleep(0.5)

#Start executing commands
for i, command in enumerate(ai_output_list):
    ai_output_list[i] = command.split(" ", 1)

print(ai_output_list)
for command in ai_output_list:
    action=command[0]
    if action=="Click":
        targetCell = command[1]
        ws.range(targetCell).select()
        time.sleep(0.2)
    elif action=="Enter":
        formula=command[1]
        ws.range(targetCell).value = formula
        time.sleep(0.2)
    elif action=="Press":
        key=command[1]
        pyautogui.press(key)
        time.sleep(0.2)

