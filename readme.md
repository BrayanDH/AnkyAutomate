# Anky Flashcard Automate


## Description

Anky Flashcard Automate is a bot designed to make flashcard generation easy and efficient. It takes data provided in an Excel file and converts it into a deck of flashcards. The data in the Excel file should be formatted with two specific columns: "question" and "answer."

## How to Use

1. Install Dependencies:
   Before using the bot, ensure you have the required dependencies installed. Run the following command in your terminal:

   ```
   pip install -r requirements.txt
   ```

2. Prepare Your Excel File:
   Format your data in an Excel file with the "question" and "answer" columns. Each row should contain a question and its corresponding answer.

   | Question   | Answer              |
   | ---------- | ------------------- |
   | Example 1  | Response Example 1  |
   | Example 2  | Response Example 2  |
   | ...        | ...                 |

3. Customize Variables:
   Modify the following variables in the script to match your specific Excel file and desired deck name:

   ```python
   excel_file = "words.xlsx"   # Replace "words.xlsx" with your Excel file name
   deck_name = "My English Words"   # Replace "My English Words" with your desired deck name
   ```

4. Run the Script:
   Open any terminal in the program's folder and execute the following command:

   ```
   python flashcard_automate.py
   ```

   The bot will process the data in the Excel file and generate the flashcard deck specified by the deck name.

Enjoy your flashcards and happy learning!

