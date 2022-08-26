import genanki
import openpyxl

# The excel file name we want use
excel_file = "words"

# Your deck name in anki
deck_name = "My English Words"

# The row we want to start with
row = 1

my_deck = genanki.Deck(
    2059400110,
    f'{deck_name}')


def get_value(book, row, column):
    wb = openpyxl.load_workbook(f'{book}.xlsx')
    ws = wb.active
    data = ws.cell(row=row, column=column).value
    wb.close()
    return data


def add_note(question_data, answer_data):
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
                'name': f'Card {question_data}',
                'qfmt': '{{Question}}',
                'afmt': '{{FrontSide}}<hr id="answer">{{Answer}}',
            },
        ], css=style)

    print(question_data, answer_data)
    print("Add values")
    my_note = genanki.Note(
        model=my_model,
        fields=[f'{question_data}', f'{answer_data}'])
    my_deck.add_note(my_note)


check_data_existence = get_value(f"{excel_file}", row, 1)
if check_data_existence == None:
    print("No have any value in the firt row, can't create the deck")
    exit()

while check_data_existence != None:

    if check_data_existence == None:
        question_data = None
        answer_data = None
        print("No have more values")
        break

    question_data = get_value(f"{excel_file}", row, 1)
    answer_data = get_value(f"{excel_file}", row, 2)
    print("Extracting values from excel file")

    if question_data != None and answer_data != None:
        add_note(question_data, answer_data)
        row += 1
        check_data_existence = get_value(f"{excel_file}", row, 1)

    elif question_data != None and answer_data == None:
        print(f"We have a problem with your notes, please add secundary value in the all notes ")
        exit()

    elif question_data == None and answer_data != None:
        print(f"We have a problem with your notes, please add values in the all question cells ")
        exit()

if question_data == None and answer_data == None or row != 1:
    print("Finished")
    print("Saving deck")
    genanki.Package(my_deck).write_to_file(f'{deck_name}.apkg')
    print("Saved")
    print("Finished")
    print("Exiting")
