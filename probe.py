import genanki
import openpyxl

# The excel file name we want use
my_excel = "words"

# The initial row we want to start with
initial_line = 1

# The initial row we want to finish with
final_line = 5

# We deck name in anki
deck_name = "My English Words"

my_deck = genanki.Deck(
    2059400110,
    f'{deck_name}')


def GetValue(Book, Row, Column):
    LineC = Row
    Book = Book
    Column = Column
    wb = openpyxl.load_workbook(f'{Book}.xlsx')
    ws = wb.active
    Value = ws.cell(row=LineC, column=Column).value
    wb.close()
    return Value


    


for file in range(initial_line, final_line):
    print("Extracting values from excel file")
    value = GetValue(f"{my_excel}", file, 1)
    value2 = GetValue(f"{my_excel}", file, 2)

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


genanki.Package(my_deck).write_to_file(f'{deck_name}.apkg')
