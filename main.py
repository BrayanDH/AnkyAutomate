import genanki
import openpyxl
import time

# The excel file name we want use
my_excel = "words"

# Your deck name in anki
deck_name = "My English Words"

# The row we want to start with
row = 1

my_deck = genanki.Deck(
    2059400110,
    f'{deck_name}')


def get_value(Book, Row, Column):
    wb = openpyxl.load_workbook(f'{Book}.xlsx')
    ws = wb.active
    Value = ws.cell(row=Row, column=Column).value
    wb.close()
    return Value



value = get_value(f"{my_excel}", row, 1)
while value != None:
    print("Extracting values from excel file")
    value = get_value(f"{my_excel}", row, 1)
    value2 = get_value(f"{my_excel}", row, 2)
    if value != None and value2 != None:
        style = """
        .card {
         font-family: times;
         font-size: 40px;
         text-align: center;
         color: black;
         background-color: white;
        }
        """

        my_model = genanki.Model(
            1607392319,
            'Simple Model',
            fields=[
                {'name': 'Question'},
                {'name': 'Answer'},
            ],
            templates=[
                {
                    'name': f'Card {value}',
                    'qfmt': '{{Question}}',
                    'afmt': '{{FrontSide}}<hr id="answer">{{Answer}}',
                },
            ], css=style)

        print("Add values")
        print(value, value2)
        my_note = genanki.Note(
            model=my_model,
            fields=[f'{value}', f'{value2}'])
        my_deck.add_note(my_note)
        row += 1
    elif value != None and value2 == None:
        print(f"We have a problem with your notes, please add secundary value in the all notes ")
        exit()   
else:
    if value == None and value2 == None:
        print("Finished")
        print("Saving deck")
        genanki.Package(my_deck).write_to_file(f'{deck_name}.apkg')
        print("Saved")
        print("Finished")
        print("Exiting")
        exit()
