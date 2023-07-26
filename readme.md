# This bot create your own flashcads from excel archive info

The bot's main purpose is to generate flashcards from data provided in an Excel file. Users need to format the data in a specific way, with columns for "question" and "answer."

you need to have the following dependencies installed:

`pip install -r requirements.txt`

You need add your info in excel in this format.

|          |                  |
| -------- | ---------------- |
| question | answer           |
| example  | response example |

Need specific your excel file name and the deck name in this vars.

`excel_file = "words"`

`deck_name = "My English Words"`

To finish your only need run this script opening any terminal in this program folder with this code.

`python main.py`

Finally this script add in this program folder one anki clickeable archive with the .apkg extension.

enjoy.
